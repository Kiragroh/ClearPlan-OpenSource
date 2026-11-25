"""Writer contract: headers, RFC 4180 escaping, number/time formatting, pseudonymisation, append rules."""
import base64
import os
import re
import sqlite3
import tempfile
import threading
import unittest
from datetime import datetime, timedelta, timezone

import _paths
from clearplan_review_core import review_log as rl

CSHARP_SCHEMA = os.path.join(_paths.SRC, "ClearPlan.Core", "ReviewLog", "ReviewLogSchema.cs")
SCHEMA_SQL = os.path.join(_paths.REVIEW_LOG_TOOLS, "schema.sql")
TEST_KEY = bytes(range(32))  # test-only key (bytes 0..31), identical to ClearPlan.Core.Tests ReviewLogTests

# Pseudonyms written by the C# ReviewLogWriter for the synthetic example (ReviewLogTests.ExportExample).
CSHARP_VECTORS = [
    ("user", ("SYNTH\\physicist1",), "usr_b3f5a0b1bbb0ba3402f32b602af0bee1"),
    ("user", ("SYNTH\\physicist2",), "usr_f05763559410bf6df09a576a7516c4f3"),
    ("patient", ("SYNTH-PAT-001",), "pat_1ac9496a60113e14ca7a3a7368b12fce"),
    ("patient", ("SYNTH-PAT-002",), "pat_21cc8d046f04f4c900fb25516dc1131a"),
    ("course", ("SYNTH-PAT-001", "C1"), "crs_2f440b68a4f76bdff12cbefbe93f40f7"),
    ("plan", ("1.2.826.0.1.3680043.10.9999.1", "SYNTH-PAT-001", "C1", "x"), "pln_ee249122b741ceff49a619323bf82257"),
]


class Formatting(unittest.TestCase):
    def test_escape(self):
        self.assertEqual("", rl.escape(None))
        self.assertEqual("abc", rl.escape("abc"))
        self.assertEqual('"a,b"', rl.escape("a,b"))
        self.assertEqual('"a""b"', rl.escape('a"b'))
        self.assertEqual('" lead"', rl.escape(" lead"))
        self.assertEqual('"trail "', rl.escape("trail "))
        self.assertEqual('"x\r\ny"', rl.escape("x\r\ny"))
        self.assertEqual("a,\"b,c\",\r\n", rl.format_line(["a", "b,c", ""]))

    def test_round_trip_through_parser(self):
        fields = ["plain", "with,comma", 'with "quote"', "multi\r\nline", " padded ", "", "Ünïcødé"]
        self.assertEqual([fields], rl.parse_csv(rl.format_line(fields)))

    def test_numbers_match_dotnet_framework_r_format(self):
        cases = [(25.0, "25"), (0.1 + 0.2, "0.30000000000000004"), (1e-05, "1E-05"), (1e16, "1E+16"),
                 (-0.0, "0"), (0.0001, "0.0001"), (123456789012345.0, "123456789012345"), (2.5, "2.5"),
                 (float("nan"), ""), (float("inf"), ""), (None, "")]
        for value, expected in cases:
            self.assertEqual(expected, rl.number(value), value)

    def test_utc_milliseconds_and_z(self):
        cet = timezone(timedelta(hours=2))
        self.assertEqual("2026-07-03T08:15:00.999Z", rl.utc(datetime(2026, 7, 3, 10, 15, 0, 999999, tzinfo=cet)))
        with self.assertRaises(ValueError):
            rl.utc(datetime(2026, 7, 3))

    def test_invariant_number_parse(self):
        self.assertEqual(5.0, rl.try_parse_invariant_number(" 5 "))
        self.assertEqual(1e-3, rl.try_parse_invariant_number("1E-3"))
        for text in ("<= 14", "5,0", "NaN", "Infinity", "", None, "1,000"):
            self.assertIsNone(rl.try_parse_invariant_number(text), text)


class Pseudonyms(unittest.TestCase):
    def test_compatible_with_csharp_writer(self):
        ps = rl.Pseudonymizer(TEST_KEY)
        for method, args, expected in CSHARP_VECTORS:
            self.assertEqual(expected, getattr(ps, method)(*args), (method, args))

    def test_deterministic_prefixed_and_keyed(self):
        a, b = rl.Pseudonymizer(TEST_KEY), rl.Pseudonymizer(TEST_KEY)
        other = rl.Pseudonymizer(bytes(range(1, 33)))
        key = a.patient("RS-PHANTOM-001")
        self.assertRegex(key, r"^pat_[0-9a-f]{32}$")
        self.assertEqual(key, b.patient("RS-PHANTOM-001"))
        self.assertEqual(key, a.patient("  rs-phantom-001 "))  # case-insensitive, trimmed
        self.assertNotEqual(key, other.patient("RS-PHANTOM-001"))
        self.assertEqual(a.user("rs-physicist"), a.user("HOSPITAL\\RS-Physicist"))
        self.assertNotEqual(a.course("P", "Case 1"), a.course("P", "case 1"))  # course IDs are case-sensitive
        self.assertRegex(a.plan("", "P", "Case 1", "Plan/BS"), r"^pln_[0-9a-f]{32}$")
        self.assertEqual("", a.patient(""))
        self.assertEqual("", a.course("P", " "))

    def test_clear_mode_and_short_key(self):
        clear = rl.Pseudonymizer(None, True)
        self.assertEqual(("clear", "RS-1"), (clear.identifier_mode, clear.patient(" RS-1 ")))
        with self.assertRaises(ValueError):
            rl.Pseudonymizer(bytes(8))

    def test_key_file_created_once_and_reused(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "state", "review-log.key")
            first = rl.load_or_create_key(path)
            second = rl.load_or_create_key(path)
            self.assertEqual(first, second)
            self.assertEqual(32, len(first))
            with open(path, "rb") as handle:
                self.assertEqual(first, base64.b64decode(handle.read()))
            with open(os.path.join(tmp, "csharp.key"), "wb") as handle:
                handle.write(base64.b64encode(TEST_KEY))  # key file format written by the C# writer
            self.assertEqual(TEST_KEY, rl.load_or_create_key(os.path.join(tmp, "csharp.key")))

    def test_scrubber(self):
        scrubber = rl.TextScrubber(["RS-PHANTOM-001", "Phantom", "ab"])
        self.assertEqual("[redacted] of [redacted] x", scrubber.scrub("rs-phantom-001 of PHANTOM x"))
        self.assertEqual("ab", scrubber.scrub("ab"))  # terms shorter than 3 characters are ignored
        self.assertEqual(2000, len(rl.TextScrubber([]).scrub("x" * 5000)))


class Headers(unittest.TestCase):
    def test_columns_equal_schema_sql(self):
        conn = sqlite3.connect(":memory:")
        with open(SCHEMA_SQL, encoding="utf-8") as handle:
            conn.executescript(handle.read())
        for table, columns in rl.COLUMNS.items():
            self.assertEqual(list(columns), [r[1] for r in conn.execute("PRAGMA table_info(%s)" % table)], table)

    @unittest.skipUnless(os.path.exists(CSHARP_SCHEMA), "C# source not present")
    def test_columns_equal_csharp_schema(self):
        with open(CSHARP_SCHEMA, encoding="utf-8") as handle:
            source = handle.read()
        names = {"review": "ReviewColumns", "goal_result": "GoalResultColumns", "plan_parameter": "PlanParameterColumns",
                 "check_finding": "CheckFindingColumns", "ruleset_revision": "RulesetRevisionColumns"}
        for table, name in names.items():
            block = re.search(name + r"\s*=\s*\{(.*?)\};", source, re.S).group(1)
            self.assertEqual(list(rl.COLUMNS[table]), re.findall(r'"([a-z0-9_]+)"', block), table)
        self.assertIn("SchemaVersion = %d;" % rl.SCHEMA_VERSION, source)


def minimal_rows(review_id="00000000-0000-0000-0000-0000000000aa"):
    context = rl.ReviewContext("live", review_id)
    context.started_utc = context.finished_utc = datetime(2026, 10, 27, 10, 0, tzinfo=timezone.utc)
    content = rl.ReviewContent(check_rows=[{"check_code": "c", "category": "k", "severity": "none", "status": "ok",
                                            "observed_value": "1", "expected_value": "1", "unit": "count",
                                            "message": "a, \"quoted\" message"}])
    return rl.build_record(content, context, rl.Pseudonymizer(TEST_KEY))


class Writer(unittest.TestCase):
    def test_utf8_no_bom_crlf_header_and_append(self):
        with tempfile.TemporaryDirectory() as tmp:
            writer = rl.ReviewLogWriter(tmp)
            writer.write(minimal_rows("00000000-0000-0000-0000-0000000000a1"))
            writer.write(minimal_rows("00000000-0000-0000-0000-0000000000a2"))
            for table in rl.COLUMNS:
                with open(os.path.join(tmp, table + ".csv"), "rb") as handle:
                    data = handle.read()
                self.assertTrue(data.startswith(rl.header_line(table).encode("utf-8")), table)
                self.assertNotIn(b"\xef\xbb\xbf", data[:3])
                self.assertEqual(data.count(b"\n"), data.count(b"\r\n"))
            with open(os.path.join(tmp, "review.csv"), "rb") as handle:
                self.assertEqual(3, len(handle.read().split(b"\r\n")) - 1)

    def test_different_header_is_refused_without_partial_rows(self):
        with tempfile.TemporaryDirectory() as tmp:
            with open(os.path.join(tmp, "review.csv"), "wb") as handle:
                handle.write(b"review_id,mode\r\n")
            with self.assertRaises(rl.ReviewLogSchemaError):
                rl.ReviewLogWriter(tmp).write(minimal_rows())
            self.assertFalse(os.path.exists(os.path.join(tmp, "check_finding.csv")))

    def test_interrupted_last_line_is_terminated(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "check_finding.csv")
            with open(path, "wb") as handle:
                handle.write(rl.header_line("check_finding").encode("utf-8") + b"partial")
            rl.ReviewLogWriter(tmp).write(minimal_rows())
            with open(path, "rb") as handle:
                self.assertIn(b"partial\r\n00000000", handle.read())

    def test_lock_serialises_writers(self):
        with tempfile.TemporaryDirectory() as tmp:
            os.makedirs(tmp, exist_ok=True)
            held, release = threading.Event(), threading.Event()

            def hold():
                with rl._DirectoryLock(tmp, 1.0):
                    held.set()
                    release.wait(5)

            thread = threading.Thread(target=hold)
            thread.start()
            held.wait(5)
            try:
                with self.assertRaises(TimeoutError):
                    rl.ReviewLogWriter(tmp, lock_timeout_seconds=0.2).write(minimal_rows())
            finally:
                release.set()
                thread.join()
            rl.ReviewLogWriter(tmp).write(minimal_rows())

    def test_failed_review_has_empty_counts(self):
        context = rl.ReviewContext("retrospective", "00000000-0000-0000-0000-0000000000f1")
        context.run_status, context.run_message = rl.RUN_FAILED, "patient_not_found"
        rows = rl.build_record(None, context, rl.Pseudonymizer(TEST_KEY))
        review = dict(zip(rl.COLUMNS["review"], rows["review"][0]))
        for column in ("n_goals", "n_goal", "n_checks", "n_findings_open", "total_dose_gy", "fractions"):
            self.assertEqual("", review[column], column)
        self.assertEqual([], rows["goal_result"] + rows["check_finding"] + rows["plan_parameter"])


if __name__ == "__main__":
    unittest.main()
