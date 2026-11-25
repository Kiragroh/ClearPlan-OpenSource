"""Local goal (constraint) tables in CSV or JSON, plus optional ClearPlan structure-alias JSON.

CSV columns (header names case-insensitive; order free):
  template_code (optional), template_structure (required, TG-263 primary name), aliases ("|"-separated,
  optional), objective (required), comparator (required: <, <=, >, >=, =), goal (required),
  variation (optional), source_label (optional), fractions / dose_per_fraction_gy / total_dose_gy
  (optional table scope; must be identical in every row that fills them).

JSON: {"schema": "ClearPlan.goal_table.v1", "table_id": "...", "scope": {"fractions": 5,
  "dose_per_fraction_gy": 5.0}, "targets": ["PTV"], "goals": [{same keys as the CSV columns}]}
  "aliases" may be a list or a "|"-separated string.

Malformed tables raise GoalTableError; a table with errors is never partially applied.
"""

import csv
import io
import json
import os

from . import evaluator, objectives
from .review_log import number

SCHEMA = "ClearPlan.goal_table.v1"
ALIAS_SCHEMA = "ClearPlan.structure_aliases.v1"


class GoalTableError(ValueError):
    pass


class GoalDefinition(object):
    def __init__(self, row_no, template_code, template_structure, aliases, objective, comparator, goal, variation,
                 source_label):
        self.row_no = row_no
        self.template_code = template_code
        self.template_structure = template_structure
        self.aliases = aliases
        self.objective = objective          # objectives.Objective
        self.comparator = comparator
        self.goal = goal
        self.variation = variation
        self.source_label = source_label


class GoalTable(object):
    def __init__(self, table_id, goals, scope=None, targets=None, path=None):
        self.table_id = table_id
        self.goals = goals
        self.scope = scope or {}
        self.targets = list(targets or [])
        self.path = path

    def aliases_for(self, template_structure):
        names = []
        for goal in self.goals:
            if goal.template_structure == template_structure:
                for alias in goal.aliases:
                    if alias not in names:
                        names.append(alias)
        return names


def _split_aliases(value):
    if value is None:
        return []
    if isinstance(value, (list, tuple)):
        items = value
    else:
        items = str(value).split("|")
    out = []
    for item in items:
        item = str(item).strip()
        if item and item not in out:
            out.append(item)
    return out


def _number(value, label, row_no, required):
    if value is None or (isinstance(value, str) and value.strip() == ""):
        if required:
            raise GoalTableError("Row %d: %s is required." % (row_no, label))
        return None
    if isinstance(value, bool):
        raise GoalTableError("Row %d: %s must be a number." % (row_no, label))
    if isinstance(value, (int, float)):
        parsed = float(value)
    else:
        try:
            parsed = evaluator.try_parse_number(str(value))
        except ValueError:
            parsed = None
        if parsed is None:
            raise GoalTableError("Row %d: %s %r is not a strict decimal number." % (row_no, label, value))
    if parsed < 0 or parsed != parsed or parsed in (float("inf"), float("-inf")):
        raise GoalTableError("Row %d: %s must be finite and non-negative." % (row_no, label))
    return parsed


def _goal_from_mapping(raw, row_no, used_codes):
    data = dict((str(k).strip().lower(), v) for k, v in raw.items())
    structure = str(data.get("template_structure") or "").strip()
    if not structure:
        raise GoalTableError("Row %d: template_structure is required." % row_no)
    try:
        objective = objectives.parse(data.get("objective"))
    except objectives.ObjectiveError as error:
        raise GoalTableError("Row %d: %s" % (row_no, error))
    comparator = str(data.get("comparator") or "").strip()
    goal_raw = data.get("goal")
    if not comparator and isinstance(goal_raw, str):
        parsed = evaluator.try_parse_goal(goal_raw)
        if parsed:
            comparator, goal_raw = parsed[0], parsed[1]
    if comparator not in evaluator.COMPARATORS:
        raise GoalTableError("Row %d: comparator must be one of %s." % (row_no, ", ".join(evaluator.COMPARATORS)))
    goal = _number(goal_raw, "goal", row_no, True)
    variation = _number(data.get("variation"), "variation", row_no, False)
    if variation is not None:
        if comparator in ("<", "<=") and variation < goal:
            raise GoalTableError("Row %d: variation must not be stricter than the goal for '%s'." % (row_no, comparator))
        if comparator in (">", ">=") and variation > goal:
            raise GoalTableError("Row %d: variation must not be stricter than the goal for '%s'." % (row_no, comparator))
        if comparator == "=":
            raise GoalTableError("Row %d: a variation tier is undefined for '='." % row_no)
    code = str(data.get("template_code") or "").strip() or "%s-%s" % (structure, objective.canonical)
    if code in used_codes:
        raise GoalTableError("Row %d: duplicate template_code %r." % (row_no, code))
    used_codes.add(code)
    return GoalDefinition(row_no, code, structure, _split_aliases(data.get("aliases")), objective, comparator, goal,
                          variation, str(data.get("source_label") or "").strip())


def _scope_from(values, row_no, scope):
    for key in ("fractions", "dose_per_fraction_gy", "total_dose_gy"):
        raw = values.get(key)
        if raw is None or (isinstance(raw, str) and raw.strip() == ""):
            continue
        value = _number(raw, key, row_no, True)
        if key == "fractions":
            if value != int(value) or value < 1:
                raise GoalTableError("Row %d: fractions must be a positive integer." % row_no)
            value = int(value)
        if key in scope and scope[key] != value:
            raise GoalTableError("Row %d: %s differs from earlier rows; one table has one scope." % (row_no, key))
        scope[key] = value
    return scope


def load(path, table_id=None):
    path = os.fspath(path)
    with open(path, "rb") as handle:
        raw = handle.read()
    text = raw.decode("utf-8-sig")
    stem = os.path.splitext(os.path.basename(path))[0]
    used = set()
    if path.lower().endswith(".json"):
        try:
            document = json.loads(text)
        except ValueError as error:
            raise GoalTableError("Goal table JSON is malformed: %s" % error)
        if not isinstance(document, dict) or document.get("schema") != SCHEMA:
            raise GoalTableError("Goal table JSON must declare schema %r." % SCHEMA)
        rows = document.get("goals")
        if not isinstance(rows, list) or not rows:
            raise GoalTableError("Goal table JSON needs a non-empty 'goals' list.")
        goals = [_goal_from_mapping(r, i, used) for i, r in enumerate(rows, start=1)]
        scope = _scope_from(dict((k, v) for k, v in (document.get("scope") or {}).items()), 0, {})
        targets = [str(t).strip() for t in (document.get("targets") or []) if str(t).strip()]
        return GoalTable(table_id or str(document.get("table_id") or stem), goals, scope, targets, path)
    reader = csv.DictReader(io.StringIO(text))
    if reader.fieldnames is None:
        raise GoalTableError("Goal table CSV is empty.")
    goals, scope = [], {}
    for i, row in enumerate(reader, start=1):
        clean = dict((str(k).strip().lower(), (v or "").strip()) for k, v in row.items() if k is not None)
        if not any(clean.values()) or clean.get("template_structure", "").startswith("#"):
            continue
        goals.append(_goal_from_mapping(clean, i, used))
        _scope_from(clean, i, scope)
    if not goals:
        raise GoalTableError("Goal table CSV has no goal rows.")
    return GoalTable(table_id or stem, goals, scope, [], path)


def load_aliases(path):
    """ClearPlan.structure_aliases.v1 -> {canonical_name: [aliases]} (laterality fields are ignored here)."""
    if path is None:
        return {}
    with open(os.fspath(path), "rb") as handle:
        document = json.loads(handle.read().decode("utf-8-sig"))
    if not isinstance(document, dict) or document.get("schema") != ALIAS_SCHEMA:
        raise GoalTableError("Alias JSON must declare schema %r." % ALIAS_SCHEMA)
    result = {}
    for entry in document.get("structures") or []:
        canonical = str(entry.get("canonical_name") or "").strip()
        if not canonical:
            raise GoalTableError("Alias entry without canonical_name.")
        if canonical in result:
            raise GoalTableError("Duplicate alias entry for %r." % canonical)
        result[canonical] = _split_aliases(entry.get("aliases"))
    return result


def scope_text(scope):
    if not scope:
        return ""
    parts = []
    if "fractions" in scope:
        parts.append("%d fx" % scope["fractions"])
    if "dose_per_fraction_gy" in scope:
        parts.append("%s Gy/fx" % number(scope["dose_per_fraction_gy"]))
    if "total_dose_gy" in scope:
        parts.append("%s Gy total" % number(scope["total_dose_gy"]))
    return ", ".join(parts)
