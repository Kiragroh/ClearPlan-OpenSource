"""Optional, identifier-free HTML summary of exactly the rows that were logged (not a clinical report)."""

import html
import os

BANNER = "UNVALIDATED EXAMPLE - NOT FOR CLINICAL USE"
_STATUS_COLOR = {"goal": "#2A9D8F", "ok": "#2A9D8F", "variation": "#EDA129", "not_met": "#D1495B",
                 "fail": "#D1495B", "not_evaluated": "#64748B", "info": "#4C78A8"}


def _cell(value):
    return "<td>%s</td>" % html.escape("" if value is None else str(value))


def _status(value):
    color = _STATUS_COLOR.get(value, "#64748B")
    return '<td><span style="color:%s;font-weight:bold">%s</span></td>' % (color, html.escape(value or ""))


def render(rows):
    """rows: the dict returned by review_log.build_record (already pseudonymised and scrubbed)."""
    from .review_log import COLUMNS, REVIEW, GOAL_RESULT, CHECK_FINDING, PLAN_PARAMETER
    review = dict(zip(COLUMNS[REVIEW], rows[REVIEW][0]))
    goal_cols = COLUMNS[GOAL_RESULT]
    check_cols = COLUMNS[CHECK_FINDING]
    param_cols = COLUMNS[PLAN_PARAMETER]
    out = ["<!DOCTYPE html><html><head><meta charset=\"utf-8\"><title>ClearPlan review summary</title>",
           "<style>body{font-family:Segoe UI,Arial,sans-serif;margin:24px}table{border-collapse:collapse;margin-bottom:18px}"
           "td,th{border:1px solid #ccc;padding:3px 8px;font-size:13px;text-align:left}"
           ".banner{background:#D1495B;color:#fff;padding:8px;font-weight:bold}</style></head><body>",
           '<div class="banner">%s</div>' % BANNER,
           "<h2>Review summary (same rows as the review log)</h2>",
           "<p>review_id %s &middot; mode %s &middot; plan_key %s &middot; goal table %s &middot; run_status %s</p>" % tuple(
               html.escape(review[k] or "") for k in ("review_id", "mode", "plan_key", "goal_table_id", "run_status")),
           "<p>Goals: %s goal, %s variation, %s not met, %s not evaluated &middot; open check findings: %s of %s</p>" % tuple(
               html.escape(review[k] or "") for k in ("n_goal", "n_variation", "n_not_met", "n_not_evaluated",
                                                       "n_findings_open", "n_checks")),
           "<h3>Clinical goals</h3><table><tr><th>Structure</th><th>Plan ROI</th><th>Mapping</th><th>Objective</th>"
           "<th>Goal</th><th>Variation</th><th>Achieved</th><th>Unit</th><th>Status</th></tr>"]
    for raw in rows[GOAL_RESULT]:
        g = dict(zip(goal_cols, raw))
        out.append("<tr>" + _cell(g["template_structure"]) + _cell(g["resolved_structure"]) + _cell(g["mapping_status"])
                   + _cell(g["objective"]) + _cell(g["comparator"] + " " + g["goal_value"]) + _cell(g["variation_value"])
                   + _cell(g["achieved_value"]) + _cell(g["unit"]) + _status(g["status"]) + "</tr>")
    out.append("</table><h3>Checks (including OK)</h3><table><tr><th>Check</th><th>Status</th><th>Observed</th>"
               "<th>Expected</th><th>Message</th></tr>")
    for raw in rows[CHECK_FINDING]:
        c = dict(zip(check_cols, raw))
        out.append("<tr>" + _cell(c["check_code"]) + _status(c["status"]) + _cell(c["observed_value"])
                   + _cell(c["expected_value"]) + _cell(c["message"]) + "</tr>")
    out.append("</table><h3>Plan parameters</h3><table><tr><th>Scope</th><th>Entity</th><th>Name</th><th>Value</th>"
               "<th>Unit</th><th>Method</th><th>Note</th></tr>")
    for raw in rows[PLAN_PARAMETER]:
        p = dict(zip(param_cols, raw))
        out.append("<tr>" + "".join(_cell(p[k]) for k in ("scope", "entity", "name", "value", "unit", "method", "note"))
                   + "</tr>")
    out.append("</table><p>%s. Values come from read-only dose queries; empty = unavailable.</p></body></html>" % BANNER)
    return "\n".join(out) + "\n"


def write(path, rows, overwrite=False):
    raw = os.fspath(path)
    if raw.startswith("\\\\") or raw.startswith("//"):
        raise ValueError("UNC output paths are not permitted for the HTML summary.")
    if not os.path.isabs(raw):
        raise ValueError("The HTML summary path must be an absolute non-UNC path.")
    if not os.path.isdir(os.path.dirname(raw)):
        raise ValueError("The local output directory must already exist.")
    with open(raw, "w" if overwrite else "x", encoding="utf-8", newline="\n") as handle:
        handle.write(render(rows))
    return raw
