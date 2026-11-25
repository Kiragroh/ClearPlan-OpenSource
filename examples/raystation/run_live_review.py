"""Live ClearPlan review of the plan currently open in RayStation (mode = live).

UNVALIDATED EXAMPLE - NOT FOR CLINICAL USE.

Reads the current Patient/Case/Plan/BeamSet read-only, applies the local goal table and checks, appends
one review to the shared review log (same five tables as ClearPlan for Eclipse) and optionally writes an
identifier-free HTML summary of the same rows. It never modifies RayStation state.

Inside RayStation: run this file as a script; it finds its run configuration via the environment
variable CLEARPLAN_RS_RUN_CONFIG or clearplan_rs_run.json next to this file.
Outside RayStation (tests): call run_live_review(config_path, get_current_fn=fake).
"""

import os
import sys
from datetime import datetime, timezone

HERE = os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else os.getcwd()
if HERE not in sys.path:
    sys.path.insert(0, HERE)

import raystation_adapter  # noqa: E402
from clearplan_review_core import checks, engine, goal_table, html_summary, review_log, run_config  # noqa: E402

ADAPTER_FILES = [os.path.join(HERE, name) for name in
                 ("raystation_adapter.py", "raystation_patient_loader.py", "run_live_review.py",
                  "run_retrospective_mining.py")]


def load_rules(config):
    table = goal_table.load(config["goal_table"], config.get("goal_table_id"))
    aliases = goal_table.load_aliases(config["structure_aliases"]) if config.get("structure_aliases") else {}
    check_config = checks.load_config(config.get("review_checks"))
    return table, aliases, check_config


def review_and_log(plan, mode, config, rules, started_utc, now_fn, review_id=None):
    """Evaluate one PlanData and append it to the log. Returns (rows, counts)."""
    table, aliases, check_config = rules
    result = engine.evaluate(plan, table, aliases, check_config, now_utc=started_utc)
    options = run_config.options(config)
    context = engine.make_context(
        plan, mode, table, engine.configuration_inputs(table, config.get("structure_aliases"), check_config),
        started_utc, None, raystation_adapter.current_user_id(), review_id=review_id,
        extra_files=[f for f in ADAPTER_FILES if os.path.exists(f)],
        run_status=review_log.RUN_PARTIAL if result.partial_reasons else review_log.RUN_OK,
        run_message="; ".join(result.partial_reasons))
    context.n_checks_excluded = result.n_checks_excluded
    context.finished_utc = now_fn()
    pseudonymizer = engine.pseudonymizer_for(config.get("key_file"), options)
    rows = review_log.build_record(result.content, context, pseudonymizer, options)
    counts = engine.write(rows, config["log_dir"], config.get("key_file"), options)
    return rows, counts


def run_live_review(config_path=None, get_current_fn=None, now_fn=None, review_id=None, overwrite_html=False):
    now_fn = now_fn or (lambda: datetime.now(timezone.utc))
    config = run_config.load(config_path, default_dir=HERE)
    rules = load_rules(config)
    started = now_fn()
    plan = raystation_adapter.read_current_plan(get_current_fn)
    rows, counts = review_and_log(plan, "live", config, rules, started, now_fn, review_id)
    html_path = None
    if config.get("html_summary_dir"):
        review_row = rows[review_log.REVIEW][0]
        html_path = os.path.join(config["html_summary_dir"], "clearplan_review_%s.html" % review_row[1])
        html_summary.write(html_path, rows, overwrite=overwrite_html)
    review = dict(zip(review_log.COLUMNS[review_log.REVIEW], rows[review_log.REVIEW][0]))
    return {"review_id": review["review_id"], "run_status": review["run_status"], "counts": counts,
            "html_summary": html_path, "n_goals": review["n_goals"], "n_findings_open": review["n_findings_open"]}


def main(argv=None):
    argv = sys.argv[1:] if argv is None else argv
    result = run_live_review(argv[0] if argv else None)
    print("ClearPlan review logged: review_id=%s run_status=%s goals=%s open findings=%s" % (
        result["review_id"], result["run_status"], result["n_goals"], result["n_findings_open"]))
    if result["html_summary"]:
        print("HTML summary: %s" % result["html_summary"])
    return 0


if __name__ == "__main__":
    sys.exit(main())
