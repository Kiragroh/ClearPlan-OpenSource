"""Same cases as ClearPlan.Core.Tests/PqmNumericEvaluatorTests.cs (Python is culture-independent)."""
import math
import unittest

import _paths  # noqa: F401
from clearplan_review_core import evaluator as ev


class EvaluatorPort(unittest.TestCase):
    def near(self, expected, actual):
        self.assertTrue(abs(expected - actual) < 1e-9, "expected %r, actual %r" % (expected, actual))

    def test_parses_decimal_parameters(self):
        for value in ("0.03", "0,03", ".03"):
            self.near(0.03, ev.try_parse_number(value))

    def test_evaluates_goals_and_variation(self):
        self.assertEqual("Not met", ev.evaluate("21 Gy", "<=20.5", ""))
        self.assertEqual("Variation", ev.evaluate("20.6 Gy", "<=20.5", "21"))
        self.assertEqual("Variation", ev.evaluate("20,6 Gy", "<=20.5", "21"))
        self.assertEqual("Goal", ev.evaluate("20.5 Gy", "<=20.5", None))
        self.assertEqual("Not met", ev.evaluate("20.5 Gy", "<20.5", None))
        self.assertEqual("Goal", ev.evaluate("21 %", ">20.5", ""))
        self.assertEqual("Goal", ev.evaluate("20.5 cc", ">=20.5", ""))
        self.assertEqual("Goal", ev.evaluate("20.5", "=20.5", ""))
        self.assertEqual("Variation", ev.evaluate("20.4", ">=20.5", "20"))

    def test_separates_measurement_from_unconfirmed_assessment(self):
        self.assertEqual("Not evaluated", ev.evaluate_for_confirmed_scope("10 Gy", "<=20", "", False))
        self.assertEqual("Not evaluated", ev.evaluate_for_confirmed_scope("30 Gy", "<=20", "", False))
        self.assertEqual("Goal", ev.evaluate_for_confirmed_scope("10 Gy", "<=20", "", True))
        self.assertEqual("Not met", ev.evaluate_for_confirmed_scope("30 Gy", "<=20", "", True))
        self.assertEqual("Not evaluated", ev.evaluate_for_confirmed_scope("Unavailable", "<=20", "", True))
        self.near(10, ev.try_parse_achieved("10 Gy"))

    def test_rejects_invalid_values(self):
        for value in (None, "", "NaN", "Infinity", "-1 Gy", "1.2.3", "1,234.5", "Unable to calculate D0.03cc",
                      "3 errors", "21 Unknown", "21 Gy extra"):
            self.assertEqual("Not evaluated", ev.evaluate(value, "<=20.5", "21"), value)
        self.assertEqual("Not evaluated", ev.evaluate("20 Gy", None, "21"))
        self.assertEqual("Not evaluated", ev.evaluate("20 Gy", "<=NaN", "21"))
        self.assertEqual("Not evaluated", ev.evaluate("20 Gy", "<=20.5", "nan"))
        self.assertEqual("Not evaluated", ev.evaluate("21 Gy", "<=20.5", "bad21"))
        self.assertEqual("Not evaluated", ev.evaluate("20 Gy", ">=-1", ""))
        self.assertEqual("Not evaluated", ev.evaluate("20 Gy", "<=20.5", "-1"))

    def test_converts_only_explicit_compatible_dose_units(self):
        cases = [(2050, "cGy", "Gy", 20.5), (20.5, "Gy", "cGy", 2050), (20.5, "Gy", "Gy", 20.5), (95, "%", "%", 95),
                 (95, "%", "Gy", None), (20.5, "Gy", "%", None), (20.5, "Unknown", "Gy", None),
                 (20.5, "Unknown", "Unknown", None), (math.nan, "Gy", "Gy", None), (math.inf, "Gy", "Gy", None),
                 (-1, "Gy", "Gy", None)]
        for value, source, target, expected in cases:
            result = ev.try_convert_dose(value, source, target)
            if expected is None:
                self.assertIsNone(result, (value, source, target))
            else:
                self.near(expected, result)

    def test_preserves_decimal_color_ratio(self):
        self.near(20.6 / 20.5 * 100, ev.try_get_goal_ratio("20.6 Gy", "<=20.5"))
        self.assertIsNone(ev.try_get_goal_ratio("Unable to calculate D0.03cc", "<=20.5"))
        self.assertIsNone(ev.try_get_goal_ratio("1 Gy", "<=0"))

    def test_complements_volume_in_its_own_unit(self):
        cases = [(30, 500, "%", 90, 70), (30, 500, "cc", 400, 470), (30, 50, "cc", 90, None), (30, 50, "%", 90, 70),
                 (101, 500, "%", None, None), (501, 500, "cc", None, None), (-1, 500, "%", None, None),
                 (math.nan, 500, "cc", None, None), (30, math.nan, "cc", None, None), (30, 500, "Gy", None, None),
                 (30, 500, "cc", -1, None)]
        for actual, organ, unit, minimum, expected in cases:
            result = ev.try_complement_volume(actual, organ, unit, minimum)
            if expected is None:
                self.assertIsNone(result, (actual, organ, unit, minimum))
            else:
                self.near(expected, result)

    def test_uses_same_unit_for_complement_color_ratio(self):
        self.near(75, ev.try_get_complement_goal_ratio("70 %", ">=60", 500.0, "%"))
        self.near(75, ev.try_get_complement_goal_ratio("470 cc", ">=460", 500.0, "cc"))
        self.assertIsNone(ev.try_get_complement_goal_ratio("70 %", ">=100", 500.0, "%"))
        self.assertIsNone(ev.try_get_complement_goal_ratio("70 %", ">=60", 500.0, "Gy"))

    def test_numeric_front_end_matches_string_evaluator(self):
        self.assertEqual("Goal", ev.evaluate_numeric(20.5, "<=", 20.5, None))
        self.assertEqual("Variation", ev.evaluate_numeric(20.6, "<=", 20.5, 21.0))
        self.assertEqual("Not evaluated", ev.evaluate_numeric(None, "<=", 20.5, 21.0))
        self.assertEqual("Not evaluated", ev.evaluate_numeric(10.0, "<=", 20.0, None, scope_confirmed=False))
        # tiny values must not turn into exponent notation (which the strict parser rejects)
        self.assertEqual("Goal", ev.evaluate_numeric(1e-07, "<=", 1.0, None))
        self.assertEqual("0.0000001", ev.format_invariant(1e-07))


if __name__ == "__main__":
    unittest.main()
