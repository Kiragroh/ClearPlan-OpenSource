"""Objective syntax, goal-table loading (CSV and JSON) and structure mapping."""
import os
import tempfile
import unittest

import _paths
from clearplan_review_core import goal_table, mapping, objectives
from clearplan_review_core.model import RoiInfo


class ObjectiveSyntax(unittest.TestCase):
    def test_canonical_forms(self):
        cases = {
            "D98%[Gy]": ("dose_at_volume", "D98%[Gy]", "Gy"),
            "D0,035cc[Gy]": ("dose_at_volume", "D0.035cc[Gy]", "Gy"),
            "D0.035 cm3 [Gy]": ("dose_at_volume", "D0.035cc[Gy]", "Gy"),
            "D2%[Gy]": ("dose_at_volume", "D2%[Gy]", "Gy"),
            "Dmax": ("dose_statistic", "Dmax[Gy]", "Gy"),
            "Max[Gy]": ("dose_statistic", "Dmax[Gy]", "Gy"),
            "Mean[cGy]": ("dose_statistic", "Dmean[cGy]", "cGy"),
            "Dmean[Gy]": ("dose_statistic", "Dmean[Gy]", "Gy"),
            "V20Gy[%]": ("volume_at_dose", "V20Gy[%]", "%"),
            "V95%[%]": ("volume_at_dose", "V95%[%]", "%"),
            "CV20Gy[cc]": ("volume_at_dose", "CV20Gy[cc]", "cm3"),
            "V2000cGy[cm3]": ("volume_at_dose", "V2000cGy[cc]", "cm3"),
        }
        for text, (kind, canonical, unit) in cases.items():
            parsed = objectives.parse(text)
            self.assertEqual((kind, canonical, unit), (parsed.kind, parsed.canonical, parsed.log_unit), text)
        self.assertTrue(objectives.parse("CV20Gy[cc]").complement)

    def test_rejects_unknown_or_unitless_forms(self):
        for text in ("", None, "D95", "V20", "D101%[Gy]", "D-1%[Gy]", "D95%[MU]", "gEUD[Gy]", "D95%Gy"):
            with self.assertRaises(objectives.ObjectiveError, msg=text):
                objectives.parse(text)


class GoalTables(unittest.TestCase):
    def test_csv_and_json_examples_are_equivalent(self):
        a = goal_table.load(os.path.join(_paths.RULES, "example_goal_table.json"))
        b = goal_table.load(os.path.join(_paths.RULES, "example_goal_table.csv"))
        key = lambda t: [(g.template_code, g.template_structure, g.aliases, g.objective.canonical, g.comparator,
                          g.goal, g.variation, g.source_label) for g in t.goals]
        self.assertEqual(key(a), key(b))
        self.assertEqual(a.scope, b.scope)
        self.assertEqual({"fractions": 5, "dose_per_fraction_gy": 5.0}, a.scope)

    def write(self, text, suffix=".csv"):
        handle = tempfile.NamedTemporaryFile("w", suffix=suffix, delete=False, encoding="utf-8")
        handle.write(text)
        handle.close()
        self.addCleanup(os.remove, handle.name)
        return handle.name

    def test_malformed_tables_are_rejected(self):
        header = "template_structure,objective,comparator,goal,variation\n"
        for body in ("PTV,D98%[Gy],>=,23.75,24\n",          # variation stricter than goal
                     "PTV,D98%[Gy],=>,23.75,\n",            # unknown comparator
                     "PTV,D98%[Gy],>=,23.75,\nPTV,D98%[Gy],>=,23.75,\n",  # duplicate template_code
                     "PTV,D98,>=,23.75,\n",                 # unitless objective
                     "PTV,D98%[Gy],>=,abc,\n",              # non-numeric goal
                     ",D98%[Gy],>=,23.75,\n"):              # missing structure
            path = self.write(header + body)
            with self.assertRaises(goal_table.GoalTableError, msg=body):
                goal_table.load(path)

    def test_comma_decimal_and_combined_goal_column(self):
        path = self.write('template_structure,objective,comparator,goal,variation\nRectum,"D0,035cc[Gy]",,"<=26,5",\n')
        table = goal_table.load(path)
        self.assertEqual(("<=", 26.5, "D0.035cc[Gy]"),
                         (table.goals[0].comparator, table.goals[0].goal, table.goals[0].objective.canonical))

    def test_inconsistent_scope_is_rejected(self):
        path = self.write("template_structure,objective,comparator,goal,fractions\nPTV,D98%[Gy],>=,23,5\n"
                          "Rectum,Dmax[Gy],<=,26,3\n")
        with self.assertRaises(goal_table.GoalTableError):
            goal_table.load(path)


class StructureMapping(unittest.TestCase):
    ROIS = [RoiInfo("PTV", "Ptv", True, 100.0), RoiInfo("Rektum", "Organ", True, 50.0),
            RoiInfo("Femur_Head_L", "Organ", True, 80.0), RoiInfo("FemoralHead_L", "Organ", True, 81.0),
            RoiInfo("Bowel", "Organ", False, None), RoiInfo("SpinalCanal", "Organ", True, 30.0)]

    def test_exact_name_ignoring_case_space_underscore(self):
        result = mapping.resolve("ptv", self.ROIS)
        self.assertEqual((mapping.RESOLVED, "PTV"), (result.status, result.roi.name))
        self.assertEqual(mapping.RESOLVED, mapping.resolve("Femur Head L", self.ROIS[:3]).status)

    def test_explicit_alias_from_table_or_alias_document(self):
        self.assertEqual("Rektum", mapping.resolve("Rectum", self.ROIS, ["Rektum"]).roi.name)
        self.assertEqual("Rektum", mapping.resolve("Rectum", self.ROIS, [], {"Rectum": ["Rektum"]}).roi.name)
        self.assertEqual(mapping.UNRESOLVED, mapping.resolve("Rectum", self.ROIS).status)

    def test_ambiguous_alias_is_never_substituted(self):
        result = mapping.resolve("Femur_Head_L", self.ROIS, ["FemoralHead_L"])
        self.assertEqual(mapping.AMBIGUOUS, result.status)
        self.assertIsNone(result.roi)
        self.assertEqual(2, len(result.candidates))

    def test_missing_and_uncontoured(self):
        self.assertEqual(mapping.UNRESOLVED, mapping.resolve("Bladder", self.ROIS).status)
        result = mapping.resolve("Bowel", self.ROIS)
        self.assertEqual(mapping.UNRESOLVED, result.status)
        self.assertIn("no contours", result.note)

    def test_no_fuzzy_or_semantic_substitution(self):
        # a canal is not the cord; a partial or cropped name is not the organ
        self.assertEqual(mapping.UNRESOLVED, mapping.resolve("SpinalCord", self.ROIS).status)
        self.assertEqual(mapping.UNRESOLVED, mapping.resolve("PTV~", self.ROIS).status)
        self.assertEqual(mapping.UNRESOLVED, mapping.resolve("PTV_Eval", self.ROIS).status)


if __name__ == "__main__":
    unittest.main()
