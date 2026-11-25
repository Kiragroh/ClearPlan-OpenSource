"""Small, configurable, TPS-independent checks. Every enabled check writes exactly one row, OK included.

Configuration (JSON, schema ClearPlan.review_checks.v1):
  ct_max_age_days       int   planning-CT age limit in days (from the examination date)
  field_name_regex      str   Python regular expression every beam name must fully match
  approved_ct_devices   list  entries with any of manufacturer / model / station / imaging_system;
                              a device is approved if all fields given in one entry match (case-insensitive)
  target_structures     list  template names of targets for target-scope descriptors (resolved via aliases)
  external_structure    str   template name of the external contour (default: the unique ROI of type External)
  disabled_checks       list  check codes to skip (counted in n_checks_excluded)
"""

import json
import math
import os
import re

SCHEMA = "ClearPlan.review_checks.v1"
CHECK_CODES = ("plan-has-dose", "fractionation-scope", "ct-age-days", "ct-device-approved", "field-naming")
DEFAULT_CONFIG = {"schema": SCHEMA, "ct_max_age_days": None, "field_name_regex": None, "approved_ct_devices": [],
                  "target_structures": [], "external_structure": None, "disabled_checks": []}


class CheckConfigError(ValueError):
    pass


def load_config(path):
    if path is None:
        return dict(DEFAULT_CONFIG, path=None)
    with open(os.fspath(path), "rb") as handle:
        try:
            document = json.loads(handle.read().decode("utf-8-sig"))
        except ValueError as error:
            raise CheckConfigError("Check configuration JSON is malformed: %s" % error)
    if not isinstance(document, dict) or document.get("schema") != SCHEMA:
        raise CheckConfigError("Check configuration must declare schema %r." % SCHEMA)
    config = dict(DEFAULT_CONFIG)
    config.update(document)
    age = config.get("ct_max_age_days")
    if age is not None and (isinstance(age, bool) or not isinstance(age, int) or age < 0):
        raise CheckConfigError("ct_max_age_days must be a non-negative integer.")
    if config.get("field_name_regex"):
        try:
            re.compile(config["field_name_regex"])
        except re.error as error:
            raise CheckConfigError("field_name_regex is invalid: %s" % error)
    for code in config.get("disabled_checks") or []:
        if code not in CHECK_CODES:
            raise CheckConfigError("Unknown check code in disabled_checks: %r" % code)
    config["path"] = os.fspath(path)
    return config


def _row(code, category, status, observed, expected, unit, message):
    severity = {"ok": "none", "info": "info", "variation": "warning", "fail": "error",
                "not_evaluated": "warning"}[status]
    return {"check_code": code, "category": category, "severity": severity, "status": status,
            "observed_value": observed, "expected_value": expected, "unit": unit, "message": message}


def _num(value):
    if value is None:
        return ""
    return ("%.6f" % value).rstrip("0").rstrip(".")


def check_plan_has_dose(plan):
    if plan.has_dose is True:
        return _row("plan-has-dose", "Dose", "ok", "Present", "Present", "boolean", "A computed plan dose is present.")
    if plan.has_dose is False:
        return _row("plan-has-dose", "Dose", "fail", "Absent", "Present", "boolean", "The plan has no computed dose.")
    return _row("plan-has-dose", "Dose", "not_evaluated", "Unknown", "Present", "boolean",
                "Dose presence could not be determined read-only.")


def scope_status(plan, scope):
    """Return (status, observed text, expected text, message, confirmed)."""
    fractions, total = plan.fractions, plan.total_dose_gy
    dpf = plan.dose_per_fraction_gy
    observed = "%s fx x %s Gy" % (fractions if fractions else "?", _num(dpf) if dpf is not None else "?")
    if plan.n_beam_sets_in_plan is not None and plan.n_beam_sets_in_plan > 1:
        return ("not_evaluated", observed, _expected(scope),
                "Plan has %d beam sets; the plan dose cannot be attributed to one fractionation." % plan.n_beam_sets_in_plan,
                False)
    if not scope:
        return ("info", observed, "", "The goal table declares no fractionation scope; goals are evaluated as configured.", True)
    if not fractions or total is None:
        return ("not_evaluated", observed, _expected(scope), "Prescription or fractionation unavailable.", False)
    mismatches = []
    if "fractions" in scope and scope["fractions"] != fractions:
        mismatches.append("fractions")
    if "dose_per_fraction_gy" in scope and abs(scope["dose_per_fraction_gy"] - dpf) > 1e-6:
        mismatches.append("dose per fraction")
    if "total_dose_gy" in scope and abs(scope["total_dose_gy"] - total) > 1e-6:
        mismatches.append("total dose")
    if mismatches:
        return ("fail", observed, _expected(scope),
                "Plan fractionation differs from the goal-table scope (%s); goals are not evaluated." % ", ".join(mismatches),
                False)
    return ("ok", observed, _expected(scope), "Plan fractionation matches the goal-table scope.", True)


def _expected(scope):
    if not scope:
        return ""
    fx = scope.get("fractions")
    dpf = scope.get("dose_per_fraction_gy")
    if dpf is None and fx and scope.get("total_dose_gy") is not None:
        dpf = scope["total_dose_gy"] / fx
    return "%s fx x %s Gy" % (fx if fx else "?", _num(dpf) if dpf is not None else "?")


def check_fractionation(plan, scope):
    status, observed, expected, message, _ = scope_status(plan, scope)
    return _row("fractionation-scope", "Prescription", status, observed, expected, "text", message)


def check_ct_age(plan, config, now_utc):
    limit = config.get("ct_max_age_days")
    expected = "<= %d" % limit if limit is not None else ""
    if limit is None:
        return _row("ct-age-days", "Metadata", "info", "", "", "count", "No planning-CT age limit is configured.")
    exam = plan.examination_datetime
    if exam is None:
        return _row("ct-age-days", "Metadata", "not_evaluated", "Unavailable", expected, "count",
                    "Planning-CT examination date unavailable.")
    seconds = (now_utc - exam).total_seconds()
    if seconds < 0:
        return _row("ct-age-days", "Metadata", "not_evaluated", "Unavailable", expected, "count",
                    "Planning-CT examination date lies in the future.")
    days = int(math.floor(seconds / 86400.0))
    if days <= limit:
        return _row("ct-age-days", "Metadata", "ok", str(days), expected, "count",
                    "Planning-CT age (days) is within the configured interval.")
    return _row("ct-age-days", "Metadata", "fail", str(days), expected, "count",
                "Planning-CT age (days) exceeds the configured interval.")


def _same(a, b):
    return (a or "").strip().lower() == (b or "").strip().lower()


def check_ct_device(plan, config):
    approved = config.get("approved_ct_devices") or []
    if not approved:
        return _row("ct-device-approved", "CT", "info", "", "", "text", "No approved CT device list is configured.")
    device = plan.ct_device
    fields = ("manufacturer", "model", "station", "imaging_system")
    for entry in approved:
        wanted = [(f, entry.get(f)) for f in fields if entry.get(f)]
        if not wanted:
            continue
        if any(getattr(device, f) in (None, "") for f, _ in wanted):
            continue
        if all(_same(getattr(device, f), value) for f, value in wanted):
            return _row("ct-device-approved", "CT", "ok", "Approved", "Approved", "text",
                        "Planning-CT device matches an approved entry.")
    needed = set(f for entry in approved for f in fields if entry.get(f))
    if any(getattr(device, f) in (None, "") for f in needed):
        return _row("ct-device-approved", "CT", "not_evaluated", "Unavailable", "Approved", "text",
                    "Planning-CT device information is missing for: %s." %
                    ", ".join(sorted(f for f in needed if getattr(device, f) in (None, ""))))
    return _row("ct-device-approved", "CT", "fail", "Not approved", "Approved", "text",
                "Planning-CT device is not on the approved list.")


def check_field_naming(plan, config):
    pattern = config.get("field_name_regex")
    if not pattern:
        return _row("field-naming", "Field naming", "info", "", "", "count", "No field-naming rule is configured.")
    if not plan.beams:
        return _row("field-naming", "Field naming", "not_evaluated", "Unavailable", "0", "count", "No treatment beams.")
    regex = re.compile(pattern)
    bad = [b for b in plan.beams if not b.name or not regex.fullmatch(b.name)]
    if not bad:
        return _row("field-naming", "Field naming", "ok", "0", "0", "count",
                    "All %d beam names match the configured rule." % len(plan.beams))
    numbers = ", ".join(str(b.number) if b.number is not None else "?" for b in bad)
    return _row("field-naming", "Field naming", "fail", str(len(bad)), "0", "count",
                "Beam names not matching the configured rule: beam number(s) %s." % numbers)


def run_checks(plan, scope, config, now_utc):
    disabled = set(config.get("disabled_checks") or [])
    rows = []
    makers = (
        ("plan-has-dose", lambda: check_plan_has_dose(plan)),
        ("fractionation-scope", lambda: check_fractionation(plan, scope)),
        ("ct-age-days", lambda: check_ct_age(plan, config, now_utc)),
        ("ct-device-approved", lambda: check_ct_device(plan, config)),
        ("field-naming", lambda: check_field_naming(plan, config)),
    )
    for code, make in makers:
        if code not in disabled:
            rows.append(make())
    return rows, len(disabled)
