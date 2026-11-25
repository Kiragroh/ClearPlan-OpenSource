"""Write the reader-facing example output from FAKE RayStation plans only (no patient data).

Usage:  python export_example_output.py <empty target dir> [<C#-written review_log dir to merge>]

Writes <target>/review_log (five CSV tables), <target>/html (summary of the live review),
<target>/review_log_raystation.sqlite + views_raystation/, and, if a C# example directory is given,
<target>/review_log_merged.sqlite + views_merged/ + import_summary_merged.json.
The key is the TEST-ONLY key (bytes 0..31) also used by the C# example, so that pseudonyms are comparable.
"""
import io
import json
import os
import sys
from contextlib import redirect_stdout
from datetime import timedelta

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _paths  # noqa: E402
import import_review_log
import test_end_to_end as e2e
import fake_raystation as fr
import run_live_review
import run_retrospective_mining


def _ids(prefix):
    counter = [0]

    def next_id():
        counter[0] += 1
        return "%s-0000-4000-8000-%012d" % (prefix, counter[0])
    return next_id


def main(argv):
    if len(argv) not in (1, 2):
        raise SystemExit(__doc__)
    target = os.path.abspath(argv[0])
    if os.path.isdir(target) and os.listdir(target):
        raise SystemExit("Target directory must be empty.")
    os.makedirs(target, exist_ok=True)
    os.environ["USERNAME"] = "rs-physicist"
    log_dir = os.path.join(target, "review_log")
    key_dir = os.path.join(target, "_test_key_not_for_production")
    key_file = os.path.join(key_dir, "review-log.key")
    e2e.write_key(key_file)
    mining_list = os.path.join(target, "mining_list.csv")
    with open(mining_list, "w", encoding="utf-8", newline="") as handle:
        handle.write("# synthetic list (fake RayStation patients)\r\npatient_id,case_name,plan_name,beam_set\r\n"
                     "RS-PHANTOM-001,Case 1,1AA_PCA_CP,1AA_PCA_CP\r\nRS-EDGE-002,Case 1,Edge,\r\n"
                     "RS-MISSING-003,Case 1,P,\r\n")
    html_dir = os.path.join(target, "html")
    os.makedirs(html_dir)
    config = e2e.write_run_config(target, log_dir, key_file, mining_list, html_dir)
    _, get_current = fr.build_phantom()
    live = run_live_review.run_live_review(config, get_current_fn=get_current, now_fn=e2e.Clock(fr.REVIEW_TIME),
                                           review_id="00000000-0000-4000-8000-00000000a001")
    phantom, _ = fr.build_phantom()
    _, db_current = fr.build_database([phantom["Patient"], e2e.edge_patient()])
    mining = run_retrospective_mining.run_retrospective_mining(
        config, get_current_fn=db_current, now_fn=e2e.Clock(fr.REVIEW_TIME + timedelta(hours=1)),
        review_id_fn=_ids("00000000"))
    summary = {"live": live, "mining": mining}
    code, single = e2e.import_dirs([log_dir], os.path.join(target, "review_log_raystation.sqlite"),
                                   os.path.join(target, "views_raystation"))
    summary["import_raystation"] = single
    if len(argv) == 2:
        code, merged = e2e.import_dirs([os.path.abspath(argv[1]), log_dir],
                                       os.path.join(target, "review_log_merged.sqlite"),
                                       os.path.join(target, "views_merged"))
        summary["import_merged"] = merged
    for key in ("import_raystation", "import_merged"):
        if key in summary:  # no absolute local paths in the published example
            summary[key]["files"] = [os.path.relpath(p, target).replace(os.sep, "/") if p.startswith(target)
                                     else "csharp_example/" + os.path.basename(p) for p in summary[key]["files"]]
    if "import_merged" in summary:
        with open(os.path.join(target, "import_summary_merged.json"), "w", encoding="utf-8") as handle:
            json.dump(summary["import_merged"], handle, indent=2, sort_keys=True)
    if live.get("html_summary"):
        live["html_summary"] = os.path.relpath(live["html_summary"], target)
    with open(os.path.join(target, "run_summary.json"), "w", encoding="utf-8") as handle:
        json.dump(summary, handle, indent=2, sort_keys=True)
    os.remove(config)
    print(json.dumps({"live": live["review_id"], "mining": mining}, sort_keys=True))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
