"""Read-only quantities, plan descriptors and checks on the fake phantom plan (simulated API only)."""
import os
import unittest
from datetime import datetime, timedelta, timezone

import _paths
import fake_raystation as fr
import raystation_adapter
from clearplan_review_core import checks, engine, goal_table

RULES = _paths.RULES
TABLE = goal_table.load(os.path.join(RULES, "example_goal_table.json"))
ALIASES = goal_table.load_aliases(os.path.join(RULES, "example_structure_aliases.json"))
CHECKS = checks.load_config(os.path.join(RULES, "example_review_checks.json"))


def review(**phantom):
    _, get_current = fr.build_phantom(**phantom)
    plan = raystation_adapter.read_current_plan(get_current)
    return plan, engine.evaluate(plan, TABLE, ALIASES, CHECKS, now_utc=fr.REVIEW_TIME)


def goals(result):
    return dict((g["template_code"], g) for g in result.content.goal_rows)


def params(result):
    return dict(((p["scope"], p["entity"], p["name"]), p) for p in result.content.parameter_rows)


def check(result, code):
    return [c for c in result.content.check_rows if c["check_code"] == code][0]


class PhantomQuantities(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.plan, cls.result = review()

    def test_goal_values_from_read_only_dose_queries(self):
        g = goals(self.result)
        expected = {"PTV-D98": (24.4, "goal"), "PTV-D2": (26.8, "variation"), "PTV-V95": (100.0, "goal"),
                    "Rectum-D0.035cc": (25.965, "goal"), "Rectum-V18": (20.0, "variation"),
                    "Rectum-CV20": (50.0 - 50.0 * (0.2 - 0.18 * 200.0 / 700.0), "goal"),
                    "Rectum-Dmean": (9.98, "goal")}
        for code, (value, status) in expected.items():
            self.assertAlmostEqual(value, g[code]["achieved_value"], places=9, msg=code)
            self.assertEqual(status, g[code]["status"], code)
        self.assertEqual("cm3", g["Rectum-CV20"]["unit"])

    def test_missing_roi_is_not_evaluated_and_has_no_value(self):
        g = goals(self.result)["Bladder-D0.035cc"]
        self.assertEqual(("unresolved", None, "not_evaluated", ""),
                         (g["mapping_status"], g["achieved_value"], g["status"], g["resolved_structure"]))

    def test_target_indices_and_mu(self):
        p = params(self.result)
        self.assertAlmostEqual(95.0 ** 2 / (100.0 * 130.0), p[("target", "PTV", "paddick_ci")]["value"], places=9)
        self.assertAlmostEqual(240.0 / 130.0, p[("target", "PTV", "gradient_index")]["value"], places=9)
        self.assertAlmostEqual((26.8 - 24.4) / 25.0, p[("target", "PTV", "homogeneity_index")]["value"], places=9)
        self.assertEqual(100.0, p[("target", "PTV", "target_volume_cm3")]["value"])
        self.assertEqual(25.0, p[("target", "PTV", "reference_dose_gy")]["value"])
        self.assertEqual(238.0, p[("plan", "", "total_mu")]["value"])
        self.assertAlmostEqual(47.6, p[("plan", "", "mu_per_gy")]["value"], places=9)
        self.assertEqual(120.0, p[("beam", "beam-1", "total_mu")]["value"])

    def test_pam_and_aperture_rows_are_unavailable_never_zero(self):
        p = params(self.result)
        for key in [("plan", "PTV", "pam"), ("beam", "beam-1", "pam"), ("plan", "", "mean_aperture_cm2"),
                    ("beam", "beam-2", "estimated_duration_s")]:
            self.assertIsNone(p[key]["value"], key)
            self.assertTrue(p[key]["note"], key)
        self.assertIn("MLC leaf geometry", p[("plan", "PTV", "pam")]["note"])

    def test_all_checks_write_a_row_including_ok(self):
        self.assertEqual(["plan-has-dose", "fractionation-scope", "ct-age-days", "ct-device-approved", "field-naming"],
                         [c["check_code"] for c in self.result.content.check_rows])
        self.assertTrue(all(c["status"] == "ok" for c in self.result.content.check_rows))
        self.assertEqual("7", check(self.result, "ct-age-days")["observed_value"])
        self.assertEqual([], self.result.partial_reasons)


class EdgeCases(unittest.TestCase):
    def test_zero_volume_roi_is_not_evaluated(self):
        _, result = review(rectum_volume=0.0)
        for code in ("Rectum-D0.035cc", "Rectum-V18", "Rectum-CV20", "Rectum-Dmean"):
            self.assertEqual("not_evaluated", goals(result)[code]["status"], code)
            self.assertIsNone(goals(result)[code]["achieved_value"], code)

    def test_ambiguous_alias_is_not_evaluated(self):
        extra = (("Rektum", "Organ", 49.0, True, fr.RECTUM),)
        _, result = review(extra_rois=extra)
        g = goals(result)["Rectum-V18"]
        self.assertEqual(("ambiguous", "", None, "not_evaluated"),
                         (g["mapping_status"], g["resolved_structure"], g["achieved_value"], g["status"]))

    def test_absolute_volume_larger_than_roi_is_unavailable(self):
        from clearplan_review_core import objectives
        plan, _ = review()
        measured = objectives.measure(objectives.parse("D60cc[Gy]"), plan.roi("Rectum"), plan.dose, 25.0)
        self.assertIsNone(measured.value)
        self.assertIn("exceeds", measured.note)
        measured = objectives.measure(objectives.parse("V95%[%]"), plan.roi("PTV"), plan.dose, None)
        self.assertIsNone(measured.value)

    def test_missing_examination_date_makes_ct_age_not_evaluated(self):
        _, result = review(exam_time=None, dicom={"EquipmentModule": fr.DICOM["EquipmentModule"]})
        self.assertEqual("not_evaluated", check(result, "ct-age-days")["status"])

    def test_old_and_future_ct(self):
        _, result = review(exam_time=fr.REVIEW_TIME - timedelta(days=45))
        self.assertEqual(("fail", "45"), (check(result, "ct-age-days")["status"],
                                          check(result, "ct-age-days")["observed_value"]))
        _, result = review(exam_time=fr.REVIEW_TIME + timedelta(days=1))
        self.assertEqual("not_evaluated", check(result, "ct-age-days")["status"])

    def test_ct_device_missing_or_not_approved(self):
        _, result = review(dicom={"StudyModule": fr.DICOM["StudyModule"]})
        self.assertEqual("not_evaluated", check(result, "ct-device-approved")["status"])
        dicom = {"EquipmentModule": {"Manufacturer": "OtherVendor", "ManufacturersModelName": "X"},
                 "StudyModule": fr.DICOM["StudyModule"]}
        _, result = review(dicom=dicom)
        self.assertEqual("fail", check(result, "ct-device-approved")["status"])

    def test_fractionation_mismatch_suppresses_assessment_but_keeps_measurement(self):
        _, result = review(fractions=3, rx_cgy=2400.0)
        self.assertEqual("fail", check(result, "fractionation-scope")["status"])
        self.assertTrue(all(g["status"] == "not_evaluated" for g in result.content.goal_rows))
        self.assertIsNotNone(goals(result)["PTV-D98"]["achieved_value"])

    def test_multiple_beam_sets_make_the_run_partial(self):
        _, result = review(n_beam_sets=2)
        self.assertEqual("not_evaluated", check(result, "fractionation-scope")["status"])
        self.assertIn("multiple_beam_sets", result.partial_reasons)
        self.assertIsNone(params(result)[("target", "PTV", "paddick_ci")]["value"])

    def test_field_naming_and_no_dose(self):
        _, result = review(beam_names=("1 AP", "Field 2"))
        row = check(result, "field-naming")
        self.assertEqual(("fail", "1"), (row["status"], row["observed_value"]))
        self.assertIn("2", row["message"])
        _, result = review(has_dose=False)
        self.assertEqual("fail", check(result, "plan-has-dose")["status"])
        self.assertTrue(all(g["achieved_value"] is None for g in result.content.goal_rows))
        self.assertIn("dose_unavailable", result.partial_reasons)

    def test_external_must_be_unique(self):
        _, result = review(external_type="Organ")
        p = params(result)[("target", "PTV", "paddick_ci")]
        self.assertIsNone(p["value"])
        self.assertIn("External", p["note"])
        self.assertIsNotNone(params(result)[("target", "PTV", "homogeneity_index")]["value"])

    def test_unconfigured_checks_are_info_and_disabled_are_counted(self):
        _, get_current = fr.build_phantom()
        plan = raystation_adapter.read_current_plan(get_current)
        config = dict(checks.DEFAULT_CONFIG, disabled_checks=["field-naming"])
        result = engine.evaluate(plan, TABLE, ALIASES, config, now_utc=fr.REVIEW_TIME)
        self.assertEqual(1, result.n_checks_excluded)
        self.assertEqual({"plan-has-dose": "ok", "fractionation-scope": "ok", "ct-age-days": "info",
                          "ct-device-approved": "info"},
                         dict((c["check_code"], c["status"]) for c in result.content.check_rows))


if __name__ == "__main__":
    unittest.main()
