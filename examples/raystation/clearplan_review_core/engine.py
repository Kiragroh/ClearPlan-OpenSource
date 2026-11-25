"""Review engine: PlanData + goal table + checks -> evaluated content -> five review-log tables.

TPS-independent. The only inputs are a PlanData (filled read-only by an adapter) and local rule files.
"""

import glob
import hashlib
import os
import re
from datetime import datetime, timezone

from . import IMPLEMENTATION, VERSION, checks, evaluator, mapping, objectives, plan_metrics, review_log


class ReviewResult(object):
    def __init__(self, content, goal_details, scope_confirmed, partial_reasons, n_checks_excluded):
        self.content = content
        self.goal_details = goal_details
        self.scope_confirmed = scope_confirmed
        self.partial_reasons = partial_reasons
        self.n_checks_excluded = n_checks_excluded


def _slug(text):
    slug = re.sub(r"[^a-z0-9.%+-]+", "-", (text or "").lower()).strip("-")
    return slug or "row"


def _resolve_targets(plan, table, alias_document, config):
    templates = list(config.get("target_structures") or []) or list(table.targets)
    results = []
    for template in templates:
        results.append(mapping.resolve(template, plan.rois, table.aliases_for(template), alias_document))
    if not templates and plan.prescription_target:
        roi = plan.roi(plan.prescription_target)
        if roi is not None and roi.has_contours is not False:
            results.append(mapping.MappingResult(roi.name, mapping.RESOLVED, roi, [roi]))
        else:
            results.append(mapping.MappingResult(plan.prescription_target, mapping.UNRESOLVED, None, [],
                                                 "Prescription target has no contoured ROI."))
    return results


def _resolve_external(plan, table, alias_document, config):
    template = config.get("external_structure")
    if template:
        result = mapping.resolve(template, plan.rois, table.aliases_for(template), alias_document)
        return result.roi
    externals = [r for r in plan.rois if (r.roi_type or "").strip().lower() == "external" and r.has_contours is not False]
    return externals[0] if len(externals) == 1 else None


def evaluate(plan, table, alias_document=None, check_config=None, now_utc=None):
    check_config = check_config or dict(checks.DEFAULT_CONFIG)
    now_utc = now_utc or datetime.now(timezone.utc)
    _, _, _, _, scope_confirmed = checks.scope_status(plan, table.scope)
    rx = plan.total_dose_gy

    goal_rows, details, used_ids = [], [], set()
    for goal in table.goals:
        result = mapping.resolve(goal.template_structure, plan.rois, goal.aliases, alias_document)
        measured = objectives.Measurement(None, result.note)
        if result.status == mapping.RESOLVED:
            measured = objectives.measure(goal.objective, result.roi, plan.dose, rx)
        status_text = evaluator.evaluate_numeric(measured.value, goal.comparator, goal.goal, goal.variation,
                                                 scope_confirmed and result.status == mapping.RESOLVED)
        stable = "pqm-" + _slug(goal.template_code)
        candidate, k = stable, 2
        while candidate in used_ids:
            candidate, k = "%s-%d" % (stable, k), k + 1
        used_ids.add(candidate)
        row = {
            "stable_id": candidate, "template_code": goal.template_code, "template_structure": goal.template_structure,
            "resolved_structure": result.roi.name if result.roi is not None else "", "mapping_status": result.status,
            "objective": goal.objective.canonical, "comparator": goal.comparator, "goal_value": goal.goal,
            "variation_value": goal.variation, "achieved_value": measured.value, "unit": goal.objective.log_unit,
            "status": evaluator.LOG_STATUS[status_text], "source_label": goal.source_label,
        }
        goal_rows.append(row)
        note = measured.note or result.note
        if not scope_confirmed and status_text == evaluator.NOT_EVALUATED and measured.value is not None:
            note = "Measured, but the goal-table scope is not confirmed for this plan."
        details.append({"row": row, "note": note, "candidates": [r.name for r in result.candidates]})

    targets = _resolve_targets(plan, table, alias_document, check_config)
    external = _resolve_external(plan, table, alias_document, check_config)
    pam_entity = targets[0].roi.name if targets and targets[0].roi is not None else ""
    parameters = plan_metrics.mu_rows(plan) + plan_metrics.unavailable_plan_rows(pam_entity)
    blocked = ""
    if plan.n_beam_sets_in_plan is not None and plan.n_beam_sets_in_plan > 1:
        blocked = "Plan has %d beam sets; Rx-based indices need a single beam set." % plan.n_beam_sets_in_plan
    for target in targets:
        rows = plan_metrics.target_rows(target.roi, external, plan.dose, rx, blocked)
        if target.roi is None:
            for r in rows:
                r["note"] = "Target '%s' %s: %s" % (target.template_structure, target.status, target.note)
        parameters.extend(rows)
    for beam in plan.beams:
        parameters.extend(plan_metrics.beam_rows(beam))

    check_rows, excluded = checks.run_checks(plan, table.scope, check_config, now_utc)
    content = review_log.ReviewContent(
        goal_rows=goal_rows, parameter_rows=parameters, check_rows=check_rows,
        machine_ids=[b.machine for b in plan.beams], techniques=[b.technique for b in plan.beams],
        mlc_models=[], total_dose_gy=plan.total_dose_gy, dose_per_fraction_gy=plan.dose_per_fraction_gy,
        fractions=plan.fractions)
    partial = list(plan.partial_reasons)
    if blocked:
        partial.append("multiple_beam_sets")
    return ReviewResult(content, details, scope_confirmed, partial, excluded)


def build_id(extra_files=()):
    """Identifies the exact code: SHA-256 over this package and the given adapter/entry files."""
    here = os.path.dirname(os.path.abspath(__file__))
    files = sorted(glob.glob(os.path.join(here, "*.py"))) + sorted(os.path.abspath(f) for f in extra_files)
    digest = hashlib.sha256()
    for path in files:
        with open(path, "rb") as handle:
            digest.update(os.path.basename(path).encode("utf-8") + b"\0" + handle.read().replace(b"\r\n", b"\n"))
    return "py-sha256:" + digest.hexdigest()[:16]


def configuration_inputs(table, alias_path, check_config):
    return [
        review_log.describe_config_file("constraints-goal-table", "constraint_catalog", table.path),
        review_log.describe_config_file("structure-aliases", "structure_aliases", alias_path),
        review_log.describe_config_file("review-checks", "check_configuration", (check_config or {}).get("path")),
    ]


def make_context(plan, mode, table, inputs, started_utc, finished_utc, user_id, review_id=None, extra_files=(),
                 run_status=None, run_message=None):
    context = review_log.ReviewContext(mode, review_id)
    context.started_utc = started_utc
    context.finished_utc = finished_utc
    context.clearplan_version = "%s/%s" % (IMPLEMENTATION, VERSION)
    context.build_id = build_id(extra_files)
    context.user_id = user_id or ""
    context.patient_id = plan.patient_id
    context.course_id = plan.course_id
    context.plan_uid = plan.plan_uid
    context.plan_id = plan.plan_id
    context.planning_item_type = "plan"
    context.plan_created_utc = plan.plan_created_utc
    context.plan_approval_status = plan.plan_approval_status or ""
    context.planning_approval_utc = plan.planning_approval_utc
    context.goal_table_id = table.table_id if table is not None else ""
    context.configuration_inputs = list(inputs)
    terms = [plan.patient_id, user_id, plan.course_id] + list(plan.patient_name_terms or [])
    context.sensitive_terms = [t for t in terms if t]
    if run_status:
        context.run_status = run_status
    if run_message:
        context.run_message = run_message
    return context


def write(rows, log_dir, key_file, options):
    writer = review_log.ReviewLogWriter(log_dir, options.lock_timeout_seconds)
    return writer.write(rows)


def pseudonymizer_for(key_file, options):
    key = None if options.log_clear_identifiers else review_log.load_or_create_key(key_file)
    return review_log.Pseudonymizer(key, options.log_clear_identifiers)
