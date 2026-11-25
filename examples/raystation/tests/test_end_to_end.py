"""Fake RayStation plan -> review-log CSV -> import_review_log.py -> SQLite -> views, plus merge with C# logs.

Set CLEARPLAN_REVIEW_LOG_EXAMPLE to the directory written by the C# writer
(ClearPlan.Core.Tests.exe ExportReviewLogExample <dir>) to run the Eclipse/RayStation merge test.
"""
import base64
import csv
import io
import json
import os
import sqlite3
import tempfile
import unittest
from contextlib import redirect_stdout
from datetime import datetime, timedelta, timezone

import _paths
import fake_raystation as fr
import import_review_log
import run_live_review
import run_retrospective_mining
from clearplan_review_core import review_log as rl

CSHARP_EXAMPLE = os.environ.get("CLEARPLAN_REVIEW_LOG_EXAMPLE")
TEST_KEY = bytes(range(32))
FORBIDDEN = ("RS-PHANTOM", "PHANTOM^", "RAYSTATION", "RS-EDGE", "rs-physicist")


class Clock(object):
    def __init__(self, start):
        self.now = start

    def __call__(self):
        self.now += timedelta(seconds=30)
        return self.now


def write_run_config(tmp, log_dir, key_file, mining_list=None, html_dir=None):
    config = {"schema": "ClearPlan.raystation_run.v1",
              "goal_table": os.path.join(_paths.RULES, "example_goal_table.json"),
              "structure_aliases": os.path.join(_paths.RULES, "example_structure_aliases.json"),
              "review_checks": os.path.join(_paths.RULES, "example_review_checks.json"),
              "log_dir": log_dir, "key_file": key_file, "html_summary_dir": html_dir, "mining_list": mining_list}
    path = os.path.join(tmp, "run.json")
    with open(path, "w", encoding="utf-8") as handle:
        json.dump(config, handle)
    return path


def write_key(path, key=TEST_KEY):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "wb") as handle:
        handle.write(base64.b64encode(key))


def edge_patient():
    """Second patient: ambiguous rectum alias, missing CT date, wrongly named field, 3 fractions."""
    objects, _ = fr.build_phantom(patient_id="RS-EDGE-002", patient_name="EDGE^RAYSTATION", plan_name="Edge",
                                  label="Edge", fractions=3, rx_cgy=2400.0, beam_names=("1 AP", "Field 2"),
                                  extra_rois=(("Rektum", "Organ", 49.0, True, fr.RECTUM),), exam_time=None,
                                  dicom={"EquipmentModule": fr.DICOM["EquipmentModule"]})
    return objects["Patient"]


def produce_raystation_log(tmp, user="rs-physicist"):
    """One live review plus a retrospective batch of four list rows (two fail)."""
    os.environ["USERNAME"] = user
    log_dir = os.path.join(tmp, "review_log")
    key_file = os.path.join(tmp, "state", "review-log.key")
    write_key(key_file)
    mining_list = os.path.join(tmp, "mining.csv")
    with open(mining_list, "w", encoding="utf-8", newline="") as handle:
        handle.write("# synthetic list\r\npatient_id,case_name,plan_name,beam_set\r\n"
                     "RS-PHANTOM-001,Case 1,1AA_PCA_CP,1AA_PCA_CP\r\nRS-EDGE-002,Case 1,Edge,\r\n"
                     "RS-MISSING-003,Case 1,P,\r\nRS-PHANTOM-001,Case 1,1AA_PCA_CP,NoSuchBeamSet\r\n")
    os.makedirs(os.path.join(tmp, "html"), exist_ok=True)
    config = write_run_config(tmp, log_dir, key_file, mining_list, os.path.join(tmp, "html"))
    _, get_current = fr.build_phantom()
    live = run_live_review.run_live_review(config, get_current_fn=get_current, now_fn=Clock(fr.REVIEW_TIME),
                                           review_id="00000000-0000-4000-8000-00000000a001")
    phantom, _ = fr.build_phantom()
    _, db_current = fr.build_database([phantom["Patient"], edge_patient()])
    mining = run_retrospective_mining.run_retrospective_mining(
        config, get_current_fn=db_current, now_fn=Clock(fr.REVIEW_TIME + timedelta(hours=1)))
    return log_dir, live, mining


def import_dirs(dirs, db, views=None):
    args = []
    for d in dirs:
        args += ["--log-dir", d]
    args += ["--db", db] + (["--export-views", views] if views else [])
    out = io.StringIO()
    with redirect_stdout(out):
        code = import_review_log.main(args)
    return code, json.loads(out.getvalue()) if code == 0 else out.getvalue()


class EndToEnd(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        cls.log_dir, cls.live, cls.mining = produce_raystation_log(cls.tmp.name)
        cls.db = os.path.join(cls.tmp.name, "review_log.sqlite")
        cls.code, cls.summary = import_dirs([cls.log_dir], cls.db, os.path.join(cls.tmp.name, "views"))
        cls.conn = sqlite3.connect(cls.db)

    @classmethod
    def tearDownClass(cls):
        cls.conn.close()
        cls.tmp.cleanup()

    def q(self, sql, *args):
        return self.conn.execute(sql, args).fetchall()

    def test_runs_and_import_succeed(self):
        self.assertEqual("ok", self.live["run_status"])
        # line numbers count the header as line 1 after '#' lines are dropped (same as the C# MiningInput)
        self.assertEqual({"targets": 4, "ok": 2, "partial": 0, "failed": 2, "write_errors": 0,
                          "lines_failed": [4, 5]}, self.mining)
        self.assertEqual(0, self.code)
        self.assertEqual(0, sum(self.summary["orphans"].values()))
        self.assertEqual({"review": 5, "goal_result": 24, "check_finding": 15, "ruleset_revision": 15},
                         dict((k, v) for k, v in self.summary["totals"].items() if k != "plan_parameter"))

    def test_modes_status_and_failed_rows(self):
        self.assertEqual([("live", "ok", 1), ("retrospective", "failed", 2), ("retrospective", "ok", 2)],
                         self.q("SELECT mode, run_status, COUNT(*) FROM review GROUP BY mode, run_status ORDER BY 1, 2"))
        self.assertEqual([("beam_set_not_found",), ("patient_not_found",)],
                         self.q("SELECT run_message FROM review WHERE run_status='failed' ORDER BY 1"))
        self.assertEqual(2, self.q("SELECT COUNT(*) FROM review WHERE run_status='failed' AND n_goals IS NULL "
                                   "AND n_checks IS NULL AND total_dose_gy IS NULL")[0][0])

    def test_live_phantom_statuses_in_sql(self):
        rows = self.q("SELECT template_code, status FROM goal_result WHERE review_id=? ORDER BY row_no",
                      self.live["review_id"])
        self.assertEqual([("PTV-D98", "goal"), ("PTV-D2", "variation"), ("PTV-V95", "goal"),
                          ("Rectum-D0.035cc", "goal"), ("Rectum-V18", "variation"), ("Rectum-CV20", "goal"),
                          ("Rectum-Dmean", "goal"), ("Bladder-D0.035cc", "not_evaluated")], rows)
        self.assertEqual([(8, 5, 2, 0, 1, 5, 0)], self.q(
            "SELECT n_goals, n_goal, n_variation, n_not_met, n_not_evaluated, n_checks, n_findings_open "
            "FROM review WHERE review_id=?", self.live["review_id"]))
        value = self.q("SELECT value FROM plan_parameter WHERE review_id=? AND name='paddick_ci'",
                       self.live["review_id"])[0][0]
        self.assertAlmostEqual(95.0 ** 2 / 13000.0, value, places=12)
        self.assertEqual(0, self.q("SELECT COUNT(*) FROM plan_parameter WHERE name='pam' AND value IS NOT NULL")[0][0])
        self.assertEqual(0, self.q("SELECT COUNT(*) FROM plan_parameter WHERE availability='unavailable' "
                                   "AND (note IS NULL OR note='')")[0][0])

    def test_edge_plan_in_sql(self):
        # a fractionation mismatch is a finding (check row fail), not a partial run
        rid = self.q("SELECT review_id FROM review WHERE fractions=3")[0][0]
        self.assertEqual([("ok", None)], self.q("SELECT run_status, run_message FROM review WHERE review_id=?", rid))
        statuses = dict(self.q("SELECT check_code, status FROM check_finding WHERE review_id=?", rid))
        self.assertEqual({"plan-has-dose": "ok", "fractionation-scope": "fail", "ct-age-days": "not_evaluated",
                          "ct-device-approved": "ok", "field-naming": "fail"}, statuses)
        self.assertEqual([("not_evaluated", 8)], self.q("SELECT status, COUNT(*) FROM goal_result WHERE review_id=? "
                                                        "GROUP BY status", rid))
        self.assertEqual(4, self.q("SELECT COUNT(*) FROM goal_result WHERE review_id=? AND mapping_status='ambiguous'",
                                   rid)[0][0])

    def test_views_and_denominators(self):
        self.assertEqual(15, self.q("SELECT SUM(n_evaluated) FROM v_recurring_findings")[0][0])
        ok = self.q("SELECT COUNT(*) FROM check_finding WHERE status='ok'")[0][0]
        self.assertGreater(ok, 0)
        self.assertEqual(1, self.q("SELECT COUNT(*) FROM v_ruleset_epoch")[0][0])
        trend = self.q("SELECT n, median FROM v_parameter_trend WHERE name='total_mu' AND scope='plan'")
        self.assertEqual([(3, 238.0)], trend)
        self.assertEqual(3, self.q("SELECT COUNT(*) FROM v_latest_review_per_plan")[0][0])

    def test_no_clear_identifiers_and_byte_format(self):
        for table in rl.COLUMNS:
            with open(os.path.join(self.log_dir, table + ".csv"), "rb") as handle:
                data = handle.read()
            self.assertTrue(data.startswith(rl.header_line(table).encode("utf-8")))
            self.assertEqual(data.count(b"\n"), data.count(b"\r\n"))
            text = data.decode("utf-8")
            for forbidden in FORBIDDEN:
                self.assertNotIn(forbidden, text, (table, forbidden))
        with open(self.live["html_summary"], encoding="utf-8") as handle:
            page = handle.read()
        self.assertIn("NOT FOR CLINICAL USE", page)
        for forbidden in FORBIDDEN:
            self.assertNotIn(forbidden, page)
        self.assertEqual(0, self.q("SELECT COUNT(*) FROM review WHERE patient_key NOT LIKE 'pat\\_%' ESCAPE '\\'")[0][0])

    def test_same_plan_links_across_live_and_retrospective(self):
        rows = self.q("SELECT mode, plan_key FROM review WHERE run_status='ok' AND fractions=5 ORDER BY mode")
        self.assertEqual(2, len(rows))
        self.assertEqual(rows[0][1], rows[1][1])

    def test_reimport_is_idempotent(self):
        code, summary = import_dirs([self.log_dir], self.db)
        self.assertEqual(0, code)
        self.assertEqual(0, sum(summary["inserted"].values()))


@unittest.skipUnless(CSHARP_EXAMPLE and os.path.isdir(CSHARP_EXAMPLE), "set CLEARPLAN_REVIEW_LOG_EXAMPLE")
class MergeWithEclipseLog(unittest.TestCase):
    """C#-written (Eclipse/ESAPI) example and RayStation-written logs in ONE database with ONE key."""

    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        cls.rs_dir, _, _ = produce_raystation_log(cls.tmp.name)
        cls.db = os.path.join(cls.tmp.name, "merged.sqlite")
        cls.code, cls.summary = import_dirs([CSHARP_EXAMPLE, cls.rs_dir], cls.db, os.path.join(cls.tmp.name, "views"))
        cls.conn = sqlite3.connect(cls.db)
        with open(os.path.join(CSHARP_EXAMPLE, "expected_counts.json"), encoding="utf-8") as handle:
            cls.expected = json.load(handle)

    @classmethod
    def tearDownClass(cls):
        cls.conn.close()
        cls.tmp.cleanup()

    def test_import_without_conflicts(self):
        self.assertEqual(0, self.code)
        self.assertEqual(0, sum(self.summary["skipped_duplicates"].values()))
        self.assertEqual(0, sum(self.summary["orphans"].values()))
        rs_rows = {}
        for table in rl.COLUMNS:
            with open(os.path.join(self.rs_dir, table + ".csv"), encoding="utf-8", newline="") as handle:
                rs_rows[table] = len(list(csv.reader(handle))) - 1
            self.assertEqual(self.expected["rows_" + table] + rs_rows[table], self.summary["totals"][table], table)

    def test_both_implementations_visible_in_views(self):
        versions = [r[0] for r in self.conn.execute("SELECT DISTINCT clearplan_version FROM review ORDER BY 1")]
        self.assertEqual(["3.2.0-synthetic", "clearplan-raystation-example/0.1.0"], versions)
        tables = [r[0] for r in self.conn.execute("SELECT DISTINCT goal_table_id FROM v_goal_status_by_table ORDER BY 1")]
        self.assertIn("Example prostate 5x5 Gy (illustrative)", tables)
        self.assertIn("Synthetic H&N 30x2Gy", tables)
        self.assertEqual(3, self.conn.execute("SELECT COUNT(*) FROM v_ruleset_epoch").fetchone()[0])
        self.assertEqual(self.summary["totals"]["check_finding"],
                         self.conn.execute("SELECT SUM(n_evaluated) FROM v_recurring_findings").fetchone()[0]
                         + self.conn.execute("SELECT COUNT(*) FROM check_finding c JOIN review r USING (review_id) "
                                             "WHERE r.run_status='failed' OR c.included=0").fetchone()[0])

    def test_shared_key_links_users_across_tps(self):
        # same login + same key file -> same usr_ pseudonym in Eclipse and RayStation rows
        with tempfile.TemporaryDirectory() as tmp:
            log_dir, _, _ = produce_raystation_log(tmp, user="SYNTH\\physicist1")
            with open(os.path.join(log_dir, "review.csv"), encoding="utf-8", newline="") as handle:
                users = set(r["user_key"] for r in csv.DictReader(handle))
        self.assertEqual({"usr_b3f5a0b1bbb0ba3402f32b602af0bee1"}, users)

    def test_python_reproduces_csharp_fingerprints(self):
        with open(os.path.join(CSHARP_EXAMPLE, "review.csv"), encoding="utf-8", newline="") as handle:
            reviews = list(csv.DictReader(handle))
        with open(os.path.join(CSHARP_EXAMPLE, "ruleset_revision.csv"), encoding="utf-8", newline="") as handle:
            revisions = list(csv.DictReader(handle))
        for review in reviews:
            inputs = [rl.ConfigurationInput(r["config_key"], r["category"], r["sha256"] or None, r["availability"])
                      for r in revisions if r["review_id"] == review["review_id"]]
            self.assertEqual(review["ruleset_sha256"], rl.ruleset_sha256(inputs, review["goal_table_sha256"]))

    def test_python_number_format_matches_csharp_files(self):
        numeric = {"review": ["total_dose_gy", "dose_per_fraction_gy"],
                   "goal_result": ["goal_value", "variation_value", "achieved_value"],
                   "plan_parameter": ["value"], "check_finding": ["observed_numeric", "expected_numeric"]}
        count = 0
        for table, columns in numeric.items():
            with open(os.path.join(CSHARP_EXAMPLE, table + ".csv"), encoding="utf-8", newline="") as handle:
                for row in csv.DictReader(handle):
                    for column in columns:
                        if row[column]:
                            count += 1
                            self.assertEqual(row[column], rl.number(float(row[column])), (table, column))
        self.assertGreater(count, 100)


if __name__ == "__main__":
    unittest.main()
