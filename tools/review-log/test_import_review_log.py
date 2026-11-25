"""End-to-end check: synthetic review-log CSV -> import_review_log.py -> SQLite -> example views.

Run:  python tools/review-log/test_import_review_log.py -v
Env:  CLEARPLAN_REVIEW_LOG_EXAMPLE=<dir written by ClearPlan.Core.Tests.exe ExportReviewLogExample <dir>>
Standard library only. Synthetic data only.
"""
import csv
import io
import json
import os
import sqlite3
import sys
import tempfile
import unittest
from contextlib import redirect_stdout

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import import_review_log  # noqa: E402

EXAMPLE = os.environ.get("CLEARPLAN_REVIEW_LOG_EXAMPLE")


@unittest.skipUnless(EXAMPLE and os.path.isdir(EXAMPLE), "set CLEARPLAN_REVIEW_LOG_EXAMPLE to an exported example")
class EndToEnd(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        with open(os.path.join(EXAMPLE, "expected_counts.json"), encoding="utf-8") as handle:
            cls.expected = json.load(handle)
        cls.tmp = tempfile.TemporaryDirectory()
        cls.db = os.path.join(cls.tmp.name, "review_log.sqlite")
        out = io.StringIO()
        with redirect_stdout(out):
            code = import_review_log.main(["--log-dir", EXAMPLE, "--db", cls.db])
        assert code == 0, out.getvalue()
        cls.summary = json.loads(out.getvalue())
        cls.conn = sqlite3.connect(cls.db)

    @classmethod
    def tearDownClass(cls):
        cls.conn.close()
        cls.tmp.cleanup()

    def scalar(self, sql, *args):
        return self.conn.execute(sql, args).fetchone()[0]

    def test_row_counts_match_writer(self):
        for table in import_review_log.TABLES:
            self.assertEqual(self.expected["rows_" + table], self.scalar("SELECT COUNT(*) FROM %s" % table), table)
        self.assertEqual(0, sum(self.summary["orphans"].values()))

    def test_reimport_is_idempotent(self):
        out = io.StringIO()
        with redirect_stdout(out):
            self.assertEqual(0, import_review_log.main(["--log-dir", EXAMPLE, "--db", self.db]))
        summary = json.loads(out.getvalue())
        self.assertEqual(0, sum(summary["inserted"].values()))
        self.assertEqual(self.expected["rows_review"], self.scalar("SELECT COUNT(*) FROM review"))

    def test_modes_and_failed_runs(self):
        for mode in ("live", "headless", "retrospective"):
            self.assertEqual(self.expected.get("reviews_" + mode, 0), self.scalar("SELECT COUNT(*) FROM review WHERE mode = ?", mode))
        self.assertEqual(self.expected["reviews_failed"], self.scalar("SELECT COUNT(*) FROM review WHERE run_status = 'failed'"))
        # A failed run carries no evaluated content: counts are NULL, not 0.
        self.assertEqual(1, self.scalar("SELECT COUNT(*) FROM review WHERE run_status = 'failed' AND n_goals IS NULL AND n_checks IS NULL"))

    def test_goal_status_view_totals(self):
        for status, column in (("goal", "n_goal"), ("variation", "n_variation"), ("not_met", "n_not_met"), ("not_evaluated", "n_not_evaluated")):
            self.assertEqual(self.expected.get("goal_" + status, 0), self.scalar("SELECT COALESCE(SUM(%s), 0) FROM v_goal_status_by_table" % column), status)
        # Review-level counters agree with the goal rows.
        self.assertEqual(self.scalar("SELECT SUM(n_goals) FROM review"), self.scalar("SELECT COUNT(*) FROM goal_result"))

    def test_recurring_findings_denominators_include_ok(self):
        total = self.scalar("SELECT SUM(n_evaluated) FROM v_recurring_findings")
        expected_total = sum(v for k, v in self.expected.items() if k.startswith("check_"))
        self.assertEqual(expected_total, total)
        self.assertEqual(self.expected.get("check_ok", 0), self.scalar("SELECT COUNT(*) FROM check_finding WHERE status = 'ok'"))
        self.assertGreater(self.expected.get("check_ok", 0), 0, "OK rows must exist so that shares have denominators")
        opened = sum(self.expected.get("check_" + s, 0) for s in ("variation", "fail", "not_evaluated"))
        self.assertEqual(opened, self.scalar("SELECT SUM(n_open) FROM v_recurring_findings"))

    def test_parameter_trend_uses_available_values_only(self):
        self.assertEqual(self.expected["parameters_available"], self.scalar("SELECT SUM(n) FROM v_parameter_trend"))
        self.assertEqual(0, self.scalar("SELECT COUNT(*) FROM plan_parameter WHERE availability = 'unavailable' AND value IS NOT NULL"))
        row = self.conn.execute("SELECT n, median, q1, q3, min_value, max_value FROM v_parameter_trend ORDER BY n DESC LIMIT 1").fetchone()
        n, median, q1, q3, lo, hi = row
        self.assertTrue(lo <= q1 <= median <= q3 <= hi, row)

    def test_rule_change_view_has_before_after(self):
        self.assertEqual(self.expected["ruleset_epochs"], self.scalar("SELECT COUNT(*) FROM v_ruleset_epoch"))
        with_previous = self.scalar("SELECT COUNT(*) FROM v_rule_change WHERE previous_ruleset_sha256 IS NOT NULL")
        self.assertGreater(with_previous, 0)
        changed = self.conn.execute(
            "SELECT COUNT(DISTINCT config_key) FROM v_ruleset_config GROUP BY config_key HAVING COUNT(DISTINCT sha256) > 1").fetchall()
        self.assertEqual(1, len(changed), "exactly one configuration document changed between the synthetic epochs")

    def test_no_clear_identifiers(self):
        for table in import_review_log.TABLES:
            with open(os.path.join(EXAMPLE, table + ".csv"), encoding="utf-8") as handle:
                text = handle.read()
            self.assertNotIn("SYNTH-PAT", text, table)
            self.assertNotIn("physicist", text, table)
        self.assertEqual(0, self.scalar("SELECT COUNT(*) FROM review WHERE patient_key NOT LIKE 'pat\\_%' ESCAPE '\\'"))


class Validation(unittest.TestCase):
    def test_header_mismatch_is_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            with open(os.path.join(tmp, "review.csv"), "w", encoding="utf-8", newline="") as handle:
                handle.write("review_id,mode\r\nx,live\r\n")
            err = io.StringIO()
            out = io.StringIO()
            with redirect_stdout(out):
                old = sys.stderr
                sys.stderr = err
                try:
                    code = import_review_log.main(["--log-dir", tmp, "--db", os.path.join(tmp, "x.sqlite")])
                finally:
                    sys.stderr = old
            self.assertEqual(1, code)
            self.assertIn("header mismatch", err.getvalue())

    def test_non_numeric_value_is_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            conn = sqlite3.connect(":memory:")
            with open(os.path.join(HERE, "schema.sql"), encoding="utf-8") as handle:
                conn.executescript(handle.read())
            columns = import_review_log.table_columns(conn, "plan_parameter")
            path = os.path.join(tmp, "plan_parameter.csv")
            with open(path, "w", encoding="utf-8", newline="") as handle:
                writer = csv.writer(handle)
                writer.writerow([c for c, _ in columns])
                writer.writerow(["r", "1", "plan", "", "pam", "0,42", "1", "m", "available", ""])
            with self.assertRaises(import_review_log.ImportError_):
                import_review_log.read_table(path, "plan_parameter", columns)


if __name__ == "__main__":
    unittest.main()
