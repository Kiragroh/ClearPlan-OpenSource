"""Illustrative read-only RayStation-to-ClearPlan snapshot adapter.

UNVALIDATED EXAMPLE — NOT FOR CLINICAL USE.

The module reads the currently opened RayStation context and writes a local
ClearPlan review-snapshot JSON file. It deliberately omits patient, case,
examination, plan, beam-set, beam, ROI, and DICOM identifiers. The resulting
file still contains confidential clinical dosimetry and must remain inside an
institution-approved local environment.

Target reference: RayStation v2025 SP2 / scripting API 17.2.0. The modern
``raystation`` packages and the legacy ``connect`` package are resolved lazily
so that this module can be imported and unit-tested outside RayStation.

This example demonstrates an adapter boundary, not commissioned equivalence
between ClearPlan constraints and RayStation clinical goals. All transcribed
PQM values, PlanCheck rows, field IDs, field names, and mappings therefore
remain explicitly ``not-evaluated``.
"""

import importlib
import json
import math
import os
from datetime import datetime, timezone
from pathlib import Path


_MODE_LABEL = "UNVALIDATED ILLUSTRATIVE RAYSTATION EXPORT — NOT FOR CLINICAL USE"
_ROLE_ORDER = {"target": 0, "oar": 1, "external": 2, "other": 3}
_ROLE_LABEL = {
    "target": "Target",
    "oar": "Organ at risk",
    "external": "External",
    "other": "Other structure",
}
_COLORS = (
    "#D1495B",
    "#2A9D8F",
    "#EDA129",
    "#6A4C93",
    "#4C78A8",
    "#64748B",
)


def _resolve_get_current():
    """Resolve RayStation's context accessor without importing it at load time."""

    errors = []
    for module_name in ("raystation", "raystation.v2025", "connect"):
        try:
            module = importlib.import_module(module_name)
        except ImportError as error:
            errors.append(error)
            continue
        get_current = getattr(module, "get_current", None)
        if callable(get_current):
            return get_current
    raise RuntimeError(
        "RayStation scripting interface was not found. Run this example inside "
        "the intended RayStation scripting context."
    ) from (errors[-1] if errors else None)


def export_current_review_snapshot(
    output_path,
    get_current_fn=None,
    generated_utc=None,
    dose_step_gy=1.0,
    overwrite=False,
):
    """Write a pseudonymized review snapshot from the current RayStation plan.

    Parameters
    ----------
    output_path:
        Existing directory plus filename, expressed as an absolute non-UNC
        path. Mapped or redirected storage still requires local review.
    get_current_fn:
        Optional dependency-injection hook for tests. In RayStation, omit it.
    generated_utc:
        Optional timezone-aware :class:`datetime` for reproducible examples.
    dose_step_gy:
        Positive finite cumulative-DVH sampling step in Gy (default 1 Gy).
    overwrite:
        Existing files are preserved unless this is explicitly ``True``.

    Returns
    -------
    pathlib.Path
        The written local JSON path.

    Notes
    -----
    The output contains confidential clinical dosimetry despite identifier
    omission. It is unvalidated and must not support clinical decisions.
    """

    destination = _validate_output_path(output_path)
    step_gy = float(dose_step_gy)
    if not math.isfinite(step_gy) or step_gy <= 0.0:
        raise ValueError("dose_step_gy must be a positive finite number.")

    get_current = get_current_fn or _resolve_get_current()
    context = _load_current_context(get_current)
    snapshot = _build_snapshot(
        context=context,
        generated_utc=_format_generated_utc(generated_utc),
        dose_step_gy=step_gy,
    )

    mode = "w" if overwrite else "x"
    with destination.open(mode, encoding="utf-8", newline="\n") as output:
        json.dump(snapshot, output, indent=2, ensure_ascii=False)
        output.write("\n")
    return destination


def _validate_output_path(output_path):
    raw_path = os.fspath(output_path)
    if raw_path.startswith("\\\\") or raw_path.startswith("//"):
        raise ValueError("UNC output paths are not permitted for this example.")
    destination = Path(raw_path)
    if not destination.is_absolute():
        raise ValueError("output_path must be an absolute non-UNC path.")
    if not destination.parent.is_dir():
        raise ValueError("The local output directory must already exist.")
    if destination.is_dir():
        raise ValueError("output_path must name a JSON file, not a directory.")
    return destination


def _format_generated_utc(value):
    if value is None:
        value = datetime.now(timezone.utc)
    if not isinstance(value, datetime):
        raise TypeError("generated_utc must be a timezone-aware datetime.")
    if value.tzinfo is None or value.utcoffset() is None:
        raise ValueError("generated_utc must include a timezone.")
    return value.astimezone(timezone.utc).isoformat(timespec="seconds")


def _load_current_context(get_current):
    context = {}
    for object_name in ("Patient", "Case", "Examination", "Plan", "BeamSet"):
        try:
            current_object = get_current(object_name)
        except Exception as error:
            raise RuntimeError(
                "Unable to obtain current RayStation {}.".format(object_name)
            ) from error
        if current_object is None:
            raise RuntimeError(
                "A current RayStation {} is required.".format(object_name)
            )
        context[object_name] = current_object
    return context


def _build_snapshot(context, generated_utc, dose_step_gy):
    case = context["Case"]
    examination = context["Examination"]
    plan = context["Plan"]
    beam_set = context["BeamSet"]
    total_dose = plan.TreatmentCourse.TotalDose

    structures = _read_structures(
        case=case,
        examination=examination,
        total_dose=total_dose,
        dose_step_gy=dose_step_gy,
    )
    raw_name_to_structure = {
        structure["raw_name"]: structure for structure in structures
    }
    plan_row = _build_plan_row(beam_set, structures)
    pqm_rows = _read_supported_goals(
        plan=plan,
        total_dose=total_dose,
        raw_name_to_structure=raw_name_to_structure,
    )
    field_rows = _build_field_rows(beam_set)

    return {
        "schemaVersion": 1,
        "scenarioId": "raystation-current-context",
        "scenarioTitle": "Illustrative RayStation review export",
        "scenarioDescription": (
            "Read-only adapter example for the current RayStation context; "
            "semantic equivalence and clinical use are not validated."
        ),
        "seed": 0,
        "synthetic": False,
        "generatedUtc": generated_utc,
        "patientDisplayLabel": "Current RayStation context",
        "planDisplayLabel": "Current RayStation plan",
        "provenanceText": (
            "Read-only illustrative extraction from the currently opened "
            "RayStation plan. Contains confidential clinical dosimetry; no "
            "patient, plan, ROI, beam, or DICOM identifiers are serialized."
        ),
        "activePlanKey": "current-plan",
        "sources": [
            {
                "stableId": "source-raystation-current",
                "sourceCode": "raystation-current-context",
                "sourceType": "RayStation v2025 SP2 illustrative adapter",
                "status": "available",
                "optional": False,
                "usedFallback": False,
                "pathDisplayLabel": "Current in-memory RayStation context",
                "message": (
                    "Read-only values were transcribed; this adapter remains "
                    "unvalidated and contains confidential clinical dosimetry."
                ),
            }
        ],
        "plans": [plan_row],
        "pqmRows": pqm_rows,
        "planCheckRows": [_not_evaluated_plancheck_row()],
        "fieldRows": field_rows,
        "structureMappings": _build_mapping_rows(structures),
        "dvhSeries": [structure["dvh"] for structure in structures],
        "report": {
            "title": "ClearPlan illustrative RayStation review",
            "subtitle": "Current plan — identifier-free display",
            "modeLabel": _MODE_LABEL,
            "watermark": _MODE_LABEL,
            "outputFileLabel": "clearplan-raystation-review.json",
            "notes": [
                "Dose is transcribed from TreatmentCourse.TotalDose in Gy.",
                "Cumulative volume is expressed as relative volume in percent.",
                "The file contains confidential clinical dosimetry.",
                "PQM and PlanCheck equivalence is not evaluated.",
            ],
        },
    }


def _read_structures(case, examination, total_dose, dose_step_gy):
    structure_set = case.PatientModel.StructureSets[examination.Name]
    candidates = []
    for roi in case.PatientModel.RegionsOfInterest:
        raw_name = str(roi.Name)
        geometry = structure_set.RoiGeometries[raw_name]
        if hasattr(geometry, "HasContours") and not geometry.HasContours():
            continue
        candidates.append(
            {
                "raw_name": raw_name,
                "role": _classify_role(roi),
                "geometry": geometry,
            }
        )

    candidates.sort(
        key=lambda candidate: (
            _ROLE_ORDER[candidate["role"]],
            candidate["raw_name"].casefold(),
        )
    )
    role_counts = {"target": 0, "oar": 0, "external": 0, "other": 0}
    for index, candidate in enumerate(candidates):
        role = candidate["role"]
        role_counts[role] += 1
        role_index = role_counts[role]
        pseudonym = "{}-{:02d}".format(role, role_index)
        display_name = "{} {:02d}".format(_ROLE_LABEL[role], role_index)
        candidate["pseudonym"] = pseudonym
        candidate["display_name"] = display_name
        candidate["dvh"] = _build_dvh_series(
            raw_name=candidate["raw_name"],
            geometry=candidate["geometry"],
            total_dose=total_dose,
            dose_step_gy=dose_step_gy,
            stable_id="dvh-{}".format(pseudonym),
            pseudonym=pseudonym,
            display_name=display_name,
            role=role,
            color_hex=_COLORS[index % len(_COLORS)],
        )
    return candidates


def _classify_role(roi):
    roi_type = str(getattr(roi, "Type", "")).strip().lower()
    if roi_type in {"gtv", "ctv", "ptv", "target"}:
        return "target"
    if roi_type == "external":
        return "external"
    if roi_type in {"organ", "oar", "organatrisk"}:
        return "oar"
    return "other"


def _build_dvh_series(
    raw_name,
    geometry,
    total_dose,
    dose_step_gy,
    stable_id,
    pseudonym,
    display_name,
    role,
    color_hex,
):
    maximum_cgy = float(
        total_dose.GetDoseStatistic(RoiName=raw_name, DoseType="Max")
    )
    if not math.isfinite(maximum_cgy) or maximum_cgy < 0.0:
        raise ValueError("RayStation returned an invalid maximum dose.")
    maximum_gy = maximum_cgy / 100.0
    final_step = int(math.ceil(maximum_gy / dose_step_gy))
    dose_values_gy = [
        round(index * dose_step_gy, 6) for index in range(final_step + 1)
    ]
    dose_values_cgy = [dose_gy * 100.0 for dose_gy in dose_values_gy]
    relative_volumes = list(
        total_dose.GetRelativeVolumeAtDoseValues(
            RoiName=raw_name,
            DoseValues=dose_values_cgy,
        )
    )
    if len(relative_volumes) != len(dose_values_gy):
        raise RuntimeError("RayStation returned an incomplete DVH sample.")

    points = []
    previous_percent = 100.0
    for dose_gy, relative_volume in zip(dose_values_gy, relative_volumes):
        volume_percent = float(relative_volume) * 100.0
        if not math.isfinite(volume_percent) or not 0.0 <= volume_percent <= 100.0:
            raise ValueError(
                "RayStation returned a DVH volume outside 0 to 100 percent."
            )
        if volume_percent > previous_percent + 1e-6:
            raise ValueError(
                "RayStation returned a cumulative DVH that is not monotone."
            )
        points.append(
            {
                "doseGy": dose_gy,
                "volumePercent": round(volume_percent, 3),
            }
        )
        previous_percent = volume_percent

    return {
        "stableId": stable_id,
        "structureId": pseudonym,
        "displayName": display_name,
        "role": role,
        "colorHex": color_hex,
        "lineStyle": "solid",
        "selected": role in {"target", "oar"},
        "volumeCc": round(float(geometry.GetRoiVolume()), 3),
        "points": points,
    }


def _build_plan_row(beam_set, structures):
    fraction_count = _optional_fraction_count(beam_set)
    prescription_gy = _optional_prescription_gy(beam_set)
    dose_per_fraction_gy = None
    if prescription_gy is not None and fraction_count:
        dose_per_fraction_gy = round(prescription_gy / fraction_count, 6)
    target = next(
        (
            structure["display_name"]
            for structure in structures
            if structure["role"] == "target"
        ),
        "Not identified",
    )
    return {
        "planKey": "current-plan",
        "displayLabel": "Current RayStation plan",
        "createdUtc": None,
        "dosePerFractionGy": dose_per_fraction_gy,
        "totalDoseGy": prescription_gy,
        "fractionCount": fraction_count,
        "targetDisplayLabel": target,
        "status": "not-evaluated",
    }


def _optional_fraction_count(beam_set):
    pattern = getattr(beam_set, "FractionationPattern", None)
    value = getattr(pattern, "NumberOfFractions", None)
    if value is None:
        return None
    count = int(value)
    return count if count > 0 else None


def _optional_prescription_gy(beam_set):
    prescription = getattr(beam_set, "Prescription", None)
    primary = getattr(prescription, "PrimaryDosePrescription", None)
    value = getattr(primary, "DoseValue", None)
    if value is None:
        return None
    dose_cgy = float(value)
    if not math.isfinite(dose_cgy) or dose_cgy < 0.0:
        return None
    return round(dose_cgy / 100.0, 6)


def _read_supported_goals(plan, total_dose, raw_name_to_structure):
    treatment_course = plan.TreatmentCourse
    evaluation_setup = getattr(treatment_course, "EvaluationSetup", None)
    functions = getattr(evaluation_setup, "EvaluationFunctions", ())
    rows = []
    for evaluation_function in functions:
        roi = getattr(evaluation_function, "ForRegionOfInterest", None)
        raw_name = str(getattr(roi, "Name", ""))
        structure = raw_name_to_structure.get(raw_name)
        planning_goal = getattr(evaluation_function, "PlanningGoal", None)
        goal_type = str(getattr(planning_goal, "Type", "")).split(".")[-1]
        if structure is None or goal_type not in {
            "DoseAtVolume",
            "VolumeAtDose",
            "AverageDose",
        }:
            continue
        row = _transcribe_goal(
            planning_goal=planning_goal,
            goal_type=goal_type,
            raw_name=raw_name,
            structure=structure,
            total_dose=total_dose,
            row_index=len(rows) + 1,
        )
        rows.append(row)
    return rows


def _transcribe_goal(
    planning_goal,
    goal_type,
    raw_name,
    structure,
    total_dose,
    row_index,
):
    acceptance_level = float(planning_goal.AcceptanceLevel)
    parameter_value = float(planning_goal.ParameterValue)
    comparator = _criteria_comparator(getattr(planning_goal, "Criteria", ""))

    if goal_type == "DoseAtVolume":
        relative_volume = (
            parameter_value / 100.0 if parameter_value > 1.0 else parameter_value
        )
        achieved_cgy = list(
            total_dose.GetDoseAtRelativeVolumes(
                RoiName=raw_name,
                RelativeVolumes=[relative_volume],
            )
        )[0]
        goal = acceptance_level / 100.0
        achieved = float(achieved_cgy) / 100.0
        unit = "Gy"
    elif goal_type == "VolumeAtDose":
        achieved_fraction = list(
            total_dose.GetRelativeVolumeAtDoseValues(
                RoiName=raw_name,
                DoseValues=[parameter_value],
            )
        )[0]
        goal = (
            acceptance_level * 100.0
            if abs(acceptance_level) <= 1.0
            else acceptance_level
        )
        achieved = float(achieved_fraction) * 100.0
        unit = "%"
    else:
        achieved_cgy = total_dose.GetDoseStatistic(
            RoiName=raw_name,
            DoseType="Average",
        )
        goal = acceptance_level / 100.0
        achieved = float(achieved_cgy) / 100.0
        unit = "Gy"

    pseudonym = structure["pseudonym"]
    return {
        "stableId": "pqm-{:03d}".format(row_index),
        "templateCode": "raystation-goal-{:03d}".format(row_index),
        "templateStructure": structure["display_name"],
        "resolvedStructureId": pseudonym,
        "structureOptions": [pseudonym],
        "objective": goal_type,
        "comparator": comparator,
        "goal": round(goal, 3),
        "variation": None,
        "achievedValue": round(achieved, 3),
        "unit": unit,
        "status": "not-evaluated",
        "severity": "info",
        "explanation": (
            "Numeric RayStation goal transcription only; ClearPlan constraint "
            "equivalence and acceptance semantics were not evaluated."
        ),
    }


def _criteria_comparator(criteria):
    normalized = str(criteria).replace(" ", "").lower()
    return ">=" if "atleast" in normalized else "<="


def _build_mapping_rows(structures):
    return [
        {
            "stableId": "mapping-{}".format(structure["pseudonym"]),
            "templateStructure": structure["display_name"],
            "selectedStructureId": structure["pseudonym"],
            "availableStructureIds": [structure["pseudonym"]],
            "status": "not-evaluated",
            "message": (
                "Role-based pseudonym only; clinical structure equivalence was "
                "not evaluated."
            ),
        }
        for structure in structures
    ]


def _not_evaluated_plancheck_row():
    return {
        "checkCode": "raystation-plancheck-not-implemented",
        "category": "Adapter coverage",
        "status": "not-evaluated",
        "severity": "info",
        "observedValue": "Not evaluated",
        "expectedValue": "Institution-commissioned checks",
        "unit": "text",
        "message": (
            "PlanCheck rules require institution-specific implementation and "
            "commissioning for RayStation."
        ),
    }


def _build_field_rows(beam_set):
    beams = list(getattr(beam_set, "Beams", ()))
    base_names = [_field_base_name(beam) for beam in beams]
    name_counts = {
        name.casefold(): sum(
            1 for candidate in base_names if candidate.casefold() == name.casefold()
        )
        for name in base_names
    }
    duplicate_indexes = {}
    rows = []
    for index, base_name in enumerate(base_names, start=1):
        normalized_name = base_name.casefold()
        suggested_name = base_name
        if name_counts[normalized_name] > 1:
            duplicate_index = duplicate_indexes.get(normalized_name, 0)
            suggested_name = _add_duplicate_suffix(base_name, duplicate_index)
            duplicate_indexes[normalized_name] = duplicate_index + 1
        rows.append(
            {
                "stableId": "field-{:02d}".format(index),
                "treatmentOrder": index,
                "beamNumber": index,
                "currentId": "Omitted",
                "expectedId": "Not evaluated",
                "currentName": "Omitted",
                "suggestedName": suggested_name,
                "idStatus": "not-evaluated",
                "nameStatus": "not-evaluated",
            }
        )
    return rows


def _field_base_name(beam):
    control_points = list(beam.ControlPoints)
    if not control_points:
        raise RuntimeError("A treatment beam has no control points.")
    start_angle = float(control_points[0].GantryAngle)
    stop_angle = float(control_points[-1].GantryAngle)
    is_arc = not _same_angle(start_angle, stop_angle)
    if is_arc:
        angle_part = "{}-{}".format(
            _angle_text(start_angle), _angle_text(stop_angle)
        )
    else:
        angle_part = _angle_text(start_angle)
    couch_angle = float(getattr(beam, "CouchRotationAngle", 0.0))
    table_part = "" if _same_angle(couch_angle, 0.0) else " T{}".format(
        _angle_text(couch_angle)
    )
    direction_part = (
        " {}".format(_arc_direction_token(beam.ArcRotationDirection))
        if is_arc
        else ""
    )
    return angle_part + table_part + direction_part


def _arc_direction_token(direction):
    normalized = str(direction).upper()
    if "COUNTER" in normalized or "CCW" in normalized:
        return "GUZ"
    return "UZ"


def _add_duplicate_suffix(base_name, index):
    suffix = chr(ord("a") + index) if index < 26 else str(index + 1)
    if base_name.upper().endswith((" UZ", " GUZ")):
        return base_name + suffix
    return base_name + " UZ" + suffix


def _same_angle(first, second):
    return abs(_normalize_angle(first) - _normalize_angle(second)) <= 0.5


def _angle_text(angle):
    return "{:.0f}".format(_normalize_angle(angle))


def _normalize_angle(angle):
    normalized = float(angle) % 360.0
    return normalized + 360.0 if normalized < 0.0 else normalized
