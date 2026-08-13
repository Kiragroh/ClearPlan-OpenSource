import ast
import importlib.util
import inspect
import json
import math
import tempfile
import types
import unittest
from datetime import datetime, timezone
from pathlib import Path
from unittest import mock


MODULE_PATH = Path(__file__).resolve().parent.parent / "clearplan_review_snapshot.py"
SPEC = importlib.util.spec_from_file_location("clearplan_review_snapshot", MODULE_PATH)
ADAPTER = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(ADAPTER)


class AttributeTrap:
    def __getattr__(self, name):
        raise AssertionError("Unexpected attribute access: {}".format(name))


class FakeRoi:
    def __init__(self, name, roi_type):
        self.Name = name
        self.Type = roi_type


class FakeGeometry:
    def __init__(self, volume_cc, has_contours=True):
        self._volume_cc = volume_cc
        self._has_contours = has_contours

    def HasContours(self):
        return self._has_contours

    def GetRoiVolume(self):
        return self._volume_cc


class FakeDose:
    def __init__(self):
        self.relative_volume_calls = []
        self.dose_statistic_calls = []
        self.dose_at_volume_calls = []
        self.max_cgy_by_roi = {
            "PTV_SECRET_123": 6200.0,
            "CORD_SECRET_456": 4100.0,
            "EXTERNAL_SECRET_789": 6500.0,
        }

    def GetDoseStatistic(self, RoiName, DoseType):
        self.dose_statistic_calls.append((RoiName, DoseType))
        if DoseType == "Max":
            return self.max_cgy_by_roi[RoiName]
        if DoseType == "Average":
            return {
                "PTV_SECRET_123": 5800.0,
                "CORD_SECRET_456": 2200.0,
                "EXTERNAL_SECRET_789": 1800.0,
            }[RoiName]
        raise AssertionError("Unexpected dose statistic {}".format(DoseType))

    def GetRelativeVolumeAtDoseValues(self, RoiName, DoseValues):
        self.relative_volume_calls.append((RoiName, list(DoseValues)))
        maximum = self.max_cgy_by_roi[RoiName]
        return [
            max(0.0, min(1.0, 1.0 - (float(dose_cgy) / maximum)))
            for dose_cgy in DoseValues
        ]

    def GetDoseAtRelativeVolumes(self, RoiName, RelativeVolumes):
        self.dose_at_volume_calls.append((RoiName, list(RelativeVolumes)))
        maximum = self.max_cgy_by_roi[RoiName]
        return [maximum * (1.0 - float(volume)) for volume in RelativeVolumes]


def fake_goal(roi, goal_type, criteria, acceptance_level, parameter_value):
    return types.SimpleNamespace(
        ForRegionOfInterest=roi,
        PlanningGoal=types.SimpleNamespace(
            Type=goal_type,
            Criteria=criteria,
            AcceptanceLevel=acceptance_level,
            ParameterValue=parameter_value,
        ),
    )


def fake_beam(start, stop, direction, couch):
    return types.SimpleNamespace(
        ArcRotationDirection=direction,
        CouchRotationAngle=couch,
        ControlPoints=[
            types.SimpleNamespace(GantryAngle=start),
            types.SimpleNamespace(GantryAngle=stop),
        ],
    )


class AnomalousDose(FakeDose):
    def GetRelativeVolumeAtDoseValues(self, RoiName, DoseValues):
        values = super().GetRelativeVolumeAtDoseValues(RoiName, DoseValues)
        if RoiName == "CORD_SECRET_456" and len(values) >= 4:
            values[1] = 1.2
            values[2] = 0.8
            values[3] = 0.9
        return values


def build_context(dose=None):
    patient = AttributeTrap()
    examination = types.SimpleNamespace(Name="CT_SECRET")
    target = FakeRoi("PTV_SECRET_123", "Ptv")
    organ = FakeRoi("CORD_SECRET_456", "Organ")
    external = FakeRoi("EXTERNAL_SECRET_789", "External")
    rois = [external, organ, target]
    geometries = {
        target.Name: FakeGeometry(182.5),
        organ.Name: FakeGeometry(31.25),
        external.Name: FakeGeometry(15600.0),
    }
    case = types.SimpleNamespace(
        Name="CASE_SECRET",
        PatientModel=types.SimpleNamespace(
            RegionsOfInterest=rois,
            StructureSets={
                examination.Name: types.SimpleNamespace(RoiGeometries=geometries)
            },
        ),
    )
    dose = dose or FakeDose()
    evaluation_functions = [
        fake_goal(target, "DoseAtVolume", "AtLeast", 5700.0, 0.95),
        fake_goal(organ, "VolumeAtDose", "AtMost", 0.25, 3000.0),
        fake_goal(organ, "AverageDose", "AtMost", 2600.0, 0.0),
        fake_goal(external, "UnknownGoal", "AtMost", 1.0, 1.0),
    ]
    plan = types.SimpleNamespace(
        Name="PLAN_SECRET",
        TreatmentCourse=types.SimpleNamespace(
            TotalDose=dose,
            EvaluationSetup=types.SimpleNamespace(
                EvaluationFunctions=evaluation_functions
            ),
        ),
    )
    beam_set = types.SimpleNamespace(
        DicomPlanLabel="BEAMSET_SECRET",
        FractionationPattern=types.SimpleNamespace(NumberOfFractions=30),
        Prescription=types.SimpleNamespace(
            PrimaryDosePrescription=types.SimpleNamespace(DoseValue=6000.0)
        ),
        Beams=[
            fake_beam(179.0, 181.0, "Clockwise", 30.0),
            fake_beam(179.0, 181.0, "Clockwise", 30.0),
            fake_beam(90.2, 90.2, "None", 0.0),
        ],
    )
    objects = {
        "Patient": patient,
        "Case": case,
        "Examination": examination,
        "Plan": plan,
        "BeamSet": beam_set,
    }
    calls = []

    def get_current(object_name):
        calls.append(object_name)
        return objects[object_name]

    return objects, calls, get_current, dose


class RayStationReviewSnapshotTests(unittest.TestCase):
    def export(self, directory, **overrides):
        objects, calls, get_current, dose = build_context()
        output_path = Path(directory) / "review.json"
        arguments = {
            "output_path": output_path,
            "get_current_fn": get_current,
            "generated_utc": datetime(2026, 7, 30, 8, 0, tzinfo=timezone.utc),
        }
        arguments.update(overrides)
        result = ADAPTER.export_current_review_snapshot(**arguments)
        return (
            json.loads(output_path.read_text(encoding="utf-8")),
            output_path,
            result,
            objects,
            calls,
            dose,
        )

    def test_public_function_has_documented_stable_signature(self):
        signature = inspect.signature(ADAPTER.export_current_review_snapshot)
        self.assertEqual(
            list(signature.parameters),
            [
                "output_path",
                "get_current_fn",
                "generated_utc",
                "dose_step_gy",
                "overwrite",
            ],
        )
        self.assertIsNone(signature.parameters["get_current_fn"].default)
        self.assertIsNone(signature.parameters["generated_utc"].default)
        self.assertEqual(signature.parameters["dose_step_gy"].default, 1.0)
        self.assertFalse(signature.parameters["overwrite"].default)

    def test_raystation_import_is_lazy_and_falls_back_in_documented_order(self):
        def legacy_get_current(_object_name):
            return None

        modules = {
            "connect": types.SimpleNamespace(get_current=legacy_get_current),
        }

        def import_module(name):
            if name in modules:
                return modules[name]
            raise ImportError(name)

        with mock.patch.object(
            ADAPTER.importlib, "import_module", side_effect=import_module
        ) as importer:
            self.assertIs(ADAPTER._resolve_get_current(), legacy_get_current)

        self.assertEqual(
            [call.args[0] for call in importer.call_args_list],
            ["raystation", "raystation.v2025", "connect"],
        )

    def test_reads_only_the_five_current_context_objects(self):
        with tempfile.TemporaryDirectory() as directory:
            _, _, _, _, calls, _ = self.export(directory)
        self.assertEqual(
            calls, ["Patient", "Case", "Examination", "Plan", "BeamSet"]
        )

    def test_each_missing_current_object_is_rejected(self):
        for missing_name in ["Patient", "Case", "Examination", "Plan", "BeamSet"]:
            with self.subTest(missing_name=missing_name):
                objects, _, _, _ = build_context()

                def get_current(object_name):
                    return None if object_name == missing_name else objects[object_name]

                with tempfile.TemporaryDirectory() as directory:
                    with self.assertRaisesRegex(RuntimeError, missing_name):
                        ADAPTER.export_current_review_snapshot(
                            Path(directory) / "review.json",
                            get_current_fn=get_current,
                        )

    def test_relative_output_path_is_rejected(self):
        _, _, get_current, _ = build_context()
        with self.assertRaisesRegex(ValueError, "absolute non-UNC path"):
            ADAPTER.export_current_review_snapshot(
                "review.json", get_current_fn=get_current
            )

    def test_unc_output_path_is_rejected(self):
        _, _, get_current, _ = build_context()
        with self.assertRaisesRegex(ValueError, "UNC"):
            ADAPTER.export_current_review_snapshot(
                r"\\server\share\review.json", get_current_fn=get_current
            )

    def test_non_positive_or_non_finite_dose_steps_are_rejected(self):
        _, _, get_current, _ = build_context()
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "review.json"
            for invalid_step in [0.0, -1.0, math.inf, math.nan]:
                with self.subTest(invalid_step=invalid_step):
                    with self.assertRaisesRegex(ValueError, "dose_step_gy"):
                        ADAPTER.export_current_review_snapshot(
                            output,
                            get_current_fn=get_current,
                            dose_step_gy=invalid_step,
                        )

    def test_existing_output_is_not_overwritten_by_default(self):
        _, _, get_current, _ = build_context()
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "review.json"
            output.write_text("keep", encoding="utf-8")
            with self.assertRaises(FileExistsError):
                ADAPTER.export_current_review_snapshot(
                    output, get_current_fn=get_current
                )
            self.assertEqual(output.read_text(encoding="utf-8"), "keep")

    def test_overwrite_must_be_explicit(self):
        _, _, get_current, _ = build_context()
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "review.json"
            output.write_text("old", encoding="utf-8")
            ADAPTER.export_current_review_snapshot(
                output,
                get_current_fn=get_current,
                generated_utc=datetime(
                    2026, 7, 30, 8, 0, tzinfo=timezone.utc
                ),
                overwrite=True,
            )
            self.assertEqual(json.loads(output.read_text())["schemaVersion"], 1)

    def test_snapshot_is_explicitly_non_synthetic_and_unvalidated(self):
        with tempfile.TemporaryDirectory() as directory:
            snapshot, _, _, _, _, _ = self.export(directory)
        self.assertEqual(snapshot["schemaVersion"], 1)
        self.assertFalse(snapshot["synthetic"])
        self.assertEqual(snapshot["generatedUtc"], "2026-07-30T08:00:00+00:00")
        combined = json.dumps(snapshot)
        self.assertIn("UNVALIDATED", combined)
        self.assertIn("NOT FOR CLINICAL USE", combined)
        self.assertIn("confidential clinical dosimetry", combined)

    def test_snapshot_contains_no_source_names_ids_or_uids(self):
        with tempfile.TemporaryDirectory() as directory:
            snapshot, _, _, _, _, _ = self.export(directory)
        serialized = json.dumps(snapshot)
        for forbidden in [
            "SECRET",
            "PTV_SECRET_123",
            "CORD_SECRET_456",
            "EXTERNAL_SECRET_789",
            "CT_SECRET",
            "CASE_SECRET",
            "PLAN_SECRET",
            "BEAMSET_SECRET",
        ]:
            self.assertNotIn(forbidden, serialized)

    def test_dvh_uses_total_dose_and_converts_cgy_to_gy(self):
        with tempfile.TemporaryDirectory() as directory:
            snapshot, _, _, _, _, dose = self.export(directory)
        target = next(
            series
            for series in snapshot["dvhSeries"]
            if series["role"] == "target"
        )
        self.assertEqual(
            [point["doseGy"] for point in target["points"]],
            [float(value) for value in range(63)],
        )
        roi_name, queried_cgy = dose.relative_volume_calls[0]
        self.assertEqual(roi_name, "PTV_SECRET_123")
        self.assertEqual(queried_cgy[:3], [0.0, 100.0, 200.0])
        self.assertEqual(target["points"][0]["volumePercent"], 100.0)

    def test_anomalous_dvh_volume_is_rejected_instead_of_repaired(self):
        objects, _, _, _ = build_context(dose=AnomalousDose())

        def get_current(object_name):
            return objects[object_name]

        with tempfile.TemporaryDirectory() as directory:
            with self.assertRaisesRegex(
                ValueError, "outside 0 to 100|not monotone"
            ):
                ADAPTER.export_current_review_snapshot(
                    Path(directory) / "review.json",
                    get_current_fn=get_current,
                )

    def test_structure_roles_are_pseudonymized_and_volumes_are_cm3(self):
        with tempfile.TemporaryDirectory() as directory:
            snapshot, _, _, _, _, _ = self.export(directory)
        self.assertEqual(
            [
                (series["structureId"], series["displayName"], series["role"])
                for series in snapshot["dvhSeries"]
            ],
            [
                ("target-01", "Target 01", "target"),
                ("oar-01", "Organ at risk 01", "oar"),
                ("external-01", "External 01", "external"),
            ],
        )
        self.assertEqual(
            [series["volumeCc"] for series in snapshot["dvhSeries"]],
            [182.5, 31.25, 15600.0],
        )

    def test_supported_raystation_goal_types_are_transcribed_numerically(self):
        with tempfile.TemporaryDirectory() as directory:
            snapshot, _, _, _, _, _ = self.export(directory)
        rows = snapshot["pqmRows"]
        self.assertEqual(
            [row["objective"] for row in rows],
            ["DoseAtVolume", "VolumeAtDose", "AverageDose"],
        )
        self.assertEqual(rows[0]["goal"], 57.0)
        self.assertAlmostEqual(rows[0]["achievedValue"], 3.1)
        self.assertEqual(rows[0]["unit"], "Gy")
        self.assertEqual(rows[1]["goal"], 25.0)
        self.assertAlmostEqual(rows[1]["achievedValue"], 26.829, places=3)
        self.assertEqual(rows[1]["unit"], "%")
        self.assertEqual(rows[2]["goal"], 26.0)
        self.assertEqual(rows[2]["achievedValue"], 22.0)
        self.assertEqual(rows[2]["unit"], "Gy")

    def test_all_pqm_rows_remain_not_evaluated_information(self):
        with tempfile.TemporaryDirectory() as directory:
            snapshot, _, _, _, _, _ = self.export(directory)
        self.assertTrue(snapshot["pqmRows"])
        self.assertTrue(
            all(row["status"] == "not-evaluated" for row in snapshot["pqmRows"])
        )
        self.assertTrue(
            all(row["severity"] == "info" for row in snapshot["pqmRows"])
        )
        self.assertTrue(
            all("equivalence" in row["explanation"].lower() for row in snapshot["pqmRows"])
        )

    def test_field_suggestions_use_geometry_but_omit_actual_ids_and_names(self):
        with tempfile.TemporaryDirectory() as directory:
            snapshot, _, _, _, _, _ = self.export(directory)
        rows = snapshot["fieldRows"]
        self.assertEqual(
            [row["suggestedName"] for row in rows],
            ["179-181 T30 UZa", "179-181 T30 UZb", "90"],
        )
        self.assertTrue(all(row["currentId"] == "Omitted" for row in rows))
        self.assertTrue(all(row["expectedId"] == "Not evaluated" for row in rows))
        self.assertTrue(all(row["currentName"] == "Omitted" for row in rows))
        self.assertTrue(all(row["idStatus"] == "not-evaluated" for row in rows))
        self.assertTrue(all(row["nameStatus"] == "not-evaluated" for row in rows))

    def test_plancheck_is_explicitly_not_evaluated(self):
        with tempfile.TemporaryDirectory() as directory:
            snapshot, _, _, _, _, _ = self.export(directory)
        self.assertEqual(len(snapshot["planCheckRows"]), 1)
        row = snapshot["planCheckRows"][0]
        self.assertEqual(row["status"], "not-evaluated")
        self.assertEqual(row["severity"], "info")
        self.assertIn("institution", row["message"].lower())

    def test_export_is_deterministic_for_a_fixed_context_and_timestamp(self):
        with tempfile.TemporaryDirectory() as first_directory:
            first, _, _, _, _, _ = self.export(first_directory)
        with tempfile.TemporaryDirectory() as second_directory:
            second, _, _, _, _, _ = self.export(second_directory)
        self.assertEqual(first, second)

    def test_source_has_no_raystation_write_like_calls_or_patient_database(self):
        source = MODULE_PATH.read_text(encoding="utf-8")
        self.assertNotIn("PatientDB", source)
        self.assertNotIn("FractionDose", source)
        self.assertIn("TotalDose", source)
        tree = ast.parse(source)
        forbidden_prefixes = (
            "Save",
            "Create",
            "Delete",
            "Update",
            "Set",
            "Add",
            "Remove",
            "Export",
            "Import",
        )
        forbidden_calls = []
        for node in ast.walk(tree):
            if not isinstance(node, ast.Call):
                continue
            if isinstance(node.func, ast.Attribute):
                call_name = node.func.attr
            else:
                continue
            if call_name.startswith(forbidden_prefixes):
                forbidden_calls.append(call_name)
        self.assertEqual(forbidden_calls, [])


if __name__ == "__main__":
    unittest.main()
