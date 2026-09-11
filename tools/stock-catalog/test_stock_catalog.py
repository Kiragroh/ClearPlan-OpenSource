"""Small independent extraction/semantic regression checks; no XLSX authoring."""
import copy
import hashlib
import json
from pathlib import Path
import unittest

from extract_stock_catalog import extract, match_key, parse_objective, public_catalog, source_record


class ObjectiveTests(unittest.TestCase):
    def test_volume_output_unit_is_not_input_dose_unit(self):
        result = parse_objective("V26Gy[%]", "<=40", 45)
        self.assertEqual((result["metric"], result["unit"], result["goal"]), ("V26Gy", "%", 40))

    def test_strict_comparator_and_original_variation_survive(self):
        result = parse_objective("D0.2cc[Gy]", "<26.2", "27.2")
        self.assertEqual((result["comparator"], result["goal"], result["variation"]), ("<", 26.2, 27.2))

    def test_complementary_volume_stays_lower_bound_and_absolute_volume(self):
        result = parse_objective("CV18Gy[cc]", ">=750", 475)
        self.assertEqual((result["comparator"], result["goal"], result["variation"], result["unit"]), (">=", 750, 475, "cc"))

    def test_mean_alias_has_explicit_preserved_unit(self):
        self.assertEqual(parse_objective("Mean[Gy]", "<=22.5", "23.5")["metric"], "Dmean")

    def test_unsupported_indices_and_malformed_volume_fail_closed(self):
        for metric in ["CI100%[%]", "GI50%[%]", "HI[%]", "DC5cc[Gy]", "V50Gy%[]", "D0.1cc"]:
            with self.subTest(metric=metric), self.assertRaises(ValueError):
                parse_objective(metric, "<=5", 10)

    def test_ranges_missing_units_and_numbers_are_not_repaired(self):
        for goal in ["20-25", "<=20 Gy", "", None, "<=NaN"]:
            with self.subTest(goal=goal), self.assertRaises(ValueError):
                parse_objective("Max[Gy]", goal, 30)

    def test_variation_direction_is_validated(self):
        for goal, variation in [("<=10", 9), (">=10", 11)]:
            with self.subTest(goal=goal), self.assertRaises(ValueError):
                parse_objective("D0.1cc[Gy]", goal, variation)

    def test_alias_comparison_preserves_cropped_and_combined_anatomy(self):
        self.assertEqual(match_key("Lung_R OAR"), match_key("Lung R_OAR"))
        self.assertNotEqual(match_key("Lungs-GTV"), match_key("Lungs"))
        self.assertNotEqual(match_key("Lung_R+L"), match_key("Lung_RL"))

    def test_ambiguous_source_range_is_not_invented(self):
        result = source_record("C_A_5-39", {})
        self.assertEqual(result["reference_ids"], "2")
        self.assertIn("without expansion", result["verification"])


class SnapshotTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.root = Path(__file__).resolve().parents[2]
        cls.output = cls.root / "artifacts/stock-catalog/stock2024.normalized.json"
        if not cls.output.exists():
            raise unittest.SkipTest("Run extraction to produce the ignored source-data intermediate first.")
        cls.data = json.loads(cls.output.read_text(encoding="utf-8"))

    def test_every_scoped_rule_accounted_for(self):
        data = self.data
        self.assertEqual(data["metadata"]["input_constraint_count"], 813)
        self.assertEqual(len(data["constraints"]), 495)
        self.assertEqual(sum(row["kind"] == "Constraint" for row in data["reviewRows"]), 318)

    def test_all_stock_schemas_require_confirmation(self):
        self.assertEqual(len(self.data["tables"]), 8)
        self.assertTrue(all(table["requires_confirmation"] for table in self.data["tables"]))
        conv = next(t for t in self.data["tables"] if t["table_id"] == "stock2024_conventional")
        self.assertIsNone(conv["fx_min"])
        self.assertIsNone(conv["dpf_min_gy"])

    def test_published_names_and_structure_ids_have_no_institution_branding(self):
        for table in self.data["tables"]:
            self.assertNotIn("UKE", table["display_name"])
        for structure in self.data["structures"]:
            self.assertEqual(structure["structure_id"], structure["canonical_name"])
            self.assertEqual(structure["codes"], "")

    def test_placeholder_and_special_context_rules_stay_review_only(self):
        active = self.data["constraints"]
        self.assertFalse(any(row["source"] == "MG" for row in active))
        self.assertFalse(any(row["metric"].startswith(("CI", "GI", "HI", "DC")) for row in active))
        for location in [("T_1Fx", 16), ("T_5Fx", 67), ("T_10Fx", 20), ("T_Conv", 50), ("T_Conv", 104)]:
            self.assertFalse(any((r["source_sheet"], r["source_row"]) == location for r in active))
            self.assertTrue(any((r["source_sheet"], r["source_row"]) == location for r in self.data["reviewRows"]))

    def test_dangerous_aliases_do_not_reappear(self):
        structures = {row["canonical_name"]: row for row in self.data["structures"]}
        forbidden = {"Brain": ["Normal Brain", "zNormal Brain_0"], "SpinalCord": ["Spinalkanal OAR"], "SpinalCanal": ["Myelon OAR"], "Lungs": ["Lungs-PTV"], "Liver": ["Liver-PTV"], "Kidneys": ["Renal cortex"], "BrachialPlex_L": ["PlexusBrachB"], "Rib": ["Chestwall", "Thoraxwand"]}
        for name, values in forbidden.items():
            aliases = {match_key(v) for v in structures[name]["aliases"].split("|")}
            self.assertFalse(aliases & {match_key(v) for v in values}, name)

    def test_no_ipsilateral_or_generic_target_auto_alias(self):
        self.assertFalse(any("ipsi" in json.dumps(r).lower() or r["canonical_name"] in ["Target", "Tumor", "PTV", "GTV"] for r in self.data["structures"]))

    def test_ids_unique_and_every_constraint_references_an_existing_structure(self):
        ids = {r["structure_id"] for r in self.data["structures"]}
        self.assertEqual(len(ids), len(self.data["structures"]))
        self.assertEqual(len({r["constraint_id"] for r in self.data["constraints"]}), len(self.data["constraints"]))
        self.assertTrue(all(row["structure_id"] in ids for row in self.data["constraints"]))

    def test_source_codes_legend_and_primary_reference_provenance_survive(self):
        t = next(s for s in self.data["sources"] if s["original_code"] == "T")
        self.assertIn("10.1016/j.ijrobp.2021.09.027", t["reference_text"])
        self.assertIn("xl/drawings/", t["source_location"])
        row = next(r for r in self.data["constraints"] if r["source_sheet"] == "T_8Fx" and r["source_row"] == 7)
        self.assertEqual((row["comparator"], row["goal"], row["variation"]), ("<", 26.2, 27.2))

    def test_executable_rule_preserves_raw_scalar_values(self):
        for row in self.data["constraints"]:
            parsed = parse_objective(row["raw_objective"], row["raw_goal"], row["raw_variation"])
            self.assertTrue(all(row[k] == parsed[k] for k in parsed), row["constraint_id"])

    def test_public_variant_removes_branding_without_changing_dose_or_anatomy(self):
        public = public_catalog(self.data)
        self.assertNotIn("UKE", json.dumps(public))
        self.assertIn("LOCAL2024", json.dumps(public))
        fields = ["constraint_id", "table_id", "structure_id", "structure_name", "metric", "unit", "comparator", "goal", "variation"]
        self.assertEqual([[r[f] for f in fields] for r in self.data["constraints"]], [[r[f] for f in fields] for r in public["constraints"]])
        self.assertEqual(self.data["structures"], public["structures"])


if __name__ == "__main__":
    unittest.main()
