"""Retrospective ClearPlan review mining over an explicit plan list (mode = retrospective).

UNVALIDATED EXAMPLE - NOT FOR CLINICAL USE.

Input: RFC 4180 CSV with header patient_id,case_name,plan_name[,beam_set]   ('#' lines ignored;
course_id / plan_id / beam_set_label are accepted as synonyms). Each row is opened read-only via
raystation_patient_loader.open_plan_readonly (VERIFY LOCALLY), reviewed with the same rules as a live
review and appended to the review log. Per-plan errors produce a review row with run_status = failed
and a short technical run_message; the batch continues. Nothing is ever saved in RayStation.
"""

import os
import sys
from datetime import datetime, timezone

HERE = os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else os.getcwd()
if HERE not in sys.path:
    sys.path.insert(0, HERE)

import raystation_adapter  # noqa: E402
import raystation_patient_loader  # noqa: E402
import run_live_review  # noqa: E402
from clearplan_review_core import engine, review_log, run_config  # noqa: E402
from clearplan_review_core.model import PlanData  # noqa: E402

SYNONYMS = {"patient_id": ("patient_id",), "case_name": ("case_name", "course_id"),
            "plan_name": ("plan_name", "plan_id"), "beam_set": ("beam_set", "beam_set_label")}


class MiningTarget(object):
    def __init__(self, line_number, patient_id, case_name, plan_name, beam_set):
        self.line_number = line_number
        self.patient_id = patient_id
        self.case_name = case_name
        self.plan_name = plan_name
        self.beam_set = beam_set


def parse_mining_list(text):
    rows = [r for r in review_log.parse_csv(text or "") if not (r and r[0].lstrip().startswith("#"))]
    if not rows:
        raise ValueError("The plan list is empty.")
    header = [h.strip().lower() for h in rows[0]]
    index = {}
    for key, names in SYNONYMS.items():
        found = [header.index(n) for n in names if n in header]
        index[key] = found[0] if found else -1
    for required in ("patient_id", "case_name", "plan_name"):
        if index[required] < 0:
            raise ValueError("The plan list requires a '%s' column." % required)
    targets = []
    for i, row in enumerate(rows[1:], start=2):
        def cell(key):
            j = index[key]
            return row[j].strip() if 0 <= j < len(row) else ""
        target = MiningTarget(i, cell("patient_id"), cell("case_name"), cell("plan_name"), cell("beam_set"))
        if not (target.patient_id and target.case_name and target.plan_name):
            raise ValueError("Plan list row %d lacks patient_id, case_name or plan_name." % i)
        targets.append(target)
    return targets


def _failed_row(target, code, config, rules, started, now_fn, review_id=None):
    table, aliases, check_config = rules
    stub = PlanData()
    stub.patient_id, stub.course_id = target.patient_id, target.case_name
    stub.plan_id = "%s/%s" % (target.plan_name, target.beam_set)
    options = run_config.options(config)
    context = engine.make_context(
        stub, "retrospective", table, engine.configuration_inputs(table, config.get("structure_aliases"), check_config),
        started, now_fn(), raystation_adapter.current_user_id(), review_id=review_id,
        extra_files=[f for f in run_live_review.ADAPTER_FILES if os.path.exists(f)],
        run_status=review_log.RUN_FAILED, run_message=code)
    rows = review_log.build_record(None, context, engine.pseudonymizer_for(config.get("key_file"), options), options)
    return engine.write(rows, config["log_dir"], config.get("key_file"), options)


def run_retrospective_mining(config_path=None, get_current_fn=None, now_fn=None, list_path=None,
                             validate_input_only=False, review_id_fn=None):
    """review_id_fn: optional callable returning the review_id for each list row (reproducible examples)."""
    now_fn = now_fn or (lambda: datetime.now(timezone.utc))
    review_id_fn = review_id_fn or (lambda: None)
    config = run_config.load(config_path, default_dir=HERE)
    list_path = list_path or config.get("mining_list")
    if not list_path:
        raise ValueError("No mining list configured.")
    with open(list_path, "rb") as handle:
        targets = parse_mining_list(handle.read().decode("utf-8-sig"))
    summary = {"targets": len(targets), "ok": 0, "partial": 0, "failed": 0, "write_errors": 0, "lines_failed": []}
    if validate_input_only:
        return summary
    rules = run_live_review.load_rules(config)
    get_current = get_current_fn or raystation_adapter.resolve_get_current()
    for target in targets:
        started = now_fn()
        review_id = review_id_fn()
        try:
            try:
                patient, case, plan, beam_set = raystation_patient_loader.open_plan_readonly(
                    get_current, target.patient_id, target.case_name, target.plan_name, target.beam_set)
                data = raystation_adapter.read_plan_data(patient, case, plan, beam_set)
            except raystation_patient_loader.LoadError as error:
                _failed_row(target, error.code, config, rules, started, now_fn, review_id)
                summary["failed"] += 1
                summary["lines_failed"].append(target.line_number)
                continue
            except Exception as error:  # read error: short technical code only, never the message
                _failed_row(target, "read_error:" + type(error).__name__, config, rules, started, now_fn,
                            review_id)
                summary["failed"] += 1
                summary["lines_failed"].append(target.line_number)
                continue
            rows, _ = run_live_review.review_and_log(data, "retrospective", config, rules, started, now_fn,
                                                     review_id)
            status = rows[review_log.REVIEW][0][review_log.COLUMNS[review_log.REVIEW].index("run_status")]
            summary[status] += 1
        except (IOError, OSError, TimeoutError, review_log.ReviewLogSchemaError):
            summary["write_errors"] += 1
            summary["lines_failed"].append(target.line_number)
    return summary


def main(argv=None):
    argv = sys.argv[1:] if argv is None else argv
    validate = "--validate-input-only" in argv
    args = [a for a in argv if a != "--validate-input-only"]
    summary = run_retrospective_mining(args[0] if args else None, validate_input_only=validate)
    print("ClearPlan retrospective mining: %(targets)d targets, %(ok)d ok, %(partial)d partial, "
          "%(failed)d failed, %(write_errors)d write errors" % summary)
    return 0 if summary["write_errors"] == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
