"""Read-only RayStation adapter for the ClearPlan review log.

UNVALIDATED EXAMPLE - NOT FOR CLINICAL USE. Not run or commissioned in a clinical RayStation.

This module (together with ``raystation_patient_loader.py``) is the ONLY place that touches the
RayStation scripting API. It converts the reviewed Patient / Case / Plan / BeamSet into the neutral
``clearplan_review_core.model.PlanData``; dose quantities are obtained lazily through
``RayStationDose`` (read-only dose queries on ``TreatmentCourse.TotalDose``).

Safety boundary (enforced by tests/test_readonly_boundary.py, an AST allowlist test):
no Save, Create, Delete, Set*, Add*, Remove*, Import, Export, Update, Compute, ROI algebra, dose
recomputation or attribute assignment on RayStation objects. Every RayStation call is listed in
README.md with its provenance (existing example vs. assumption to verify locally).

API provenance markers used below:
  [example]  used by the existing clearplan_review_snapshot.py example (still unvalidated)
  [verify]   assumption about the RayStation API that must be verified locally
"""

import importlib
import math
import os
from datetime import datetime, timezone

from clearplan_review_core.model import BeamInfo, CtDeviceInfo, DoseQuery, PlanData, RoiInfo

# Every RayStation method this package may call. The AST test rejects any other PascalCase call.
ALLOWED_RAYSTATION_METHODS = frozenset([
    "GetDoseStatistic",                # [example]
    "GetRelativeVolumeAtDoseValues",   # [example]
    "GetDoseAtRelativeVolumes",        # [example]
    "GetRoiVolume",                    # [example]
    "HasContours",                     # [example]
    "GetPlanningExamination",          # [verify]
    "GetExaminationDateTime",          # [verify]
    "GetAcquisitionDataFromDicom",     # [verify]
])


def resolve_get_current():
    """Resolve RayStation's context accessor lazily (same order as the existing example). [example]"""
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
    raise RuntimeError("RayStation scripting interface was not found. Run inside RayStation.") from (
        errors[-1] if errors else None)


def _finite(value):
    try:
        return value is not None and math.isfinite(float(value))
    except (TypeError, ValueError):
        return False


def _attr(obj, name):
    """Read an attribute; any failure -> None (recorded by the caller where it matters)."""
    if obj is None:
        return None
    try:
        return getattr(obj, name, None)
    except Exception:
        return None


def _text(value):
    if value is None:
        return ""
    return str(value).strip()


def _as_utc(value):
    """RayStation date/time (python or .NET DateTime, TPS local time) -> aware UTC datetime or None. [verify]"""
    if value is None:
        return None
    try:
        if isinstance(value, datetime):
            dt = value
        else:
            dt = datetime(int(value.Year), int(value.Month), int(value.Day), int(value.Hour), int(value.Minute),
                          int(value.Second), int(getattr(value, "Millisecond", 0)) * 1000)
        if dt.tzinfo is None:
            dt = dt.astimezone()  # interpret as local time of the workstation
        return dt.astimezone(timezone.utc)
    except Exception:
        return None


def _lookup(container, key):
    if container is None:
        return None
    try:
        return container[key]
    except Exception:
        return _attr(container, key)


class RayStationDose(DoseQuery):
    """Read-only dose queries on TreatmentCourse.TotalDose; converts cGy <-> Gy. [example]"""

    def __init__(self, total_dose):
        self._dose = total_dose

    def dose_at_relative_volumes(self, roi_name, relative_volumes):
        values = self._dose.GetDoseAtRelativeVolumes(RoiName=roi_name, RelativeVolumes=list(relative_volumes))
        return [float(v) / 100.0 for v in values]

    def relative_volumes_at_doses(self, roi_name, doses_gy):
        values = self._dose.GetRelativeVolumeAtDoseValues(RoiName=roi_name, DoseValues=[d * 100.0 for d in doses_gy])
        return [float(v) for v in values]

    def statistic(self, roi_name, kind):
        if kind not in ("Max", "Min", "Average"):
            raise ValueError(kind)
        return float(self._dose.GetDoseStatistic(RoiName=roi_name, DoseType=kind)) / 100.0


def _read_prescription(beam_set, plan_data):
    prescription = _attr(beam_set, "Prescription")                                   # [example]
    primary = _attr(prescription, "PrimaryDosePrescription")                         # [example]
    if primary is None:
        primary = _attr(prescription, "PrimaryPrescriptionDoseReference")            # [verify] newer API name
    dose_cgy = _attr(primary, "DoseValue")                                           # [example]
    if _finite(dose_cgy) and float(dose_cgy) > 0:
        plan_data.total_dose_gy = float(dose_cgy) / 100.0
    else:
        plan_data.partial_reasons.append("prescription_unavailable")
    target = _attr(_attr(primary, "OnStructure"), "Name")                            # [verify]
    plan_data.prescription_target = _text(target) or None
    fractions = _attr(_attr(beam_set, "FractionationPattern"), "NumberOfFractions")  # [example]
    try:
        count = int(fractions) if fractions is not None else None
    except (TypeError, ValueError):
        count = None
    plan_data.fractions = count if count and count > 0 else None
    if plan_data.fractions is None:
        plan_data.partial_reasons.append("fractions_unavailable")


def _read_rois(case, examination, plan_data):
    patient_model = _attr(case, "PatientModel")                                      # [example]
    structure_set = None
    try:
        structure_set = patient_model.StructureSets[examination.Name]                # [example]
    except Exception:
        plan_data.partial_reasons.append("structure_set_unavailable")
    errors = 0
    for roi in list(_attr(patient_model, "RegionsOfInterest") or []):                # [example]
        name = _text(_attr(roi, "Name"))                                             # [example]
        info = RoiInfo(name, _text(_attr(roi, "Type")) or None)                      # [example]
        geometry = None
        if structure_set is not None:
            try:
                geometry = structure_set.RoiGeometries[name]                         # [example]
            except Exception:
                geometry = None
        if geometry is not None:
            try:
                info.has_contours = bool(geometry.HasContours())                     # [example]
            except Exception:
                info.has_contours = None
            if info.has_contours is not False:
                try:
                    volume = float(geometry.GetRoiVolume())                          # [example] cm3
                    info.volume_cc = volume if math.isfinite(volume) and volume >= 0 else None
                except Exception:
                    errors += 1
        else:
            info.has_contours = None
        plan_data.rois.append(info)
    if errors:
        plan_data.partial_reasons.append("roi_volume_unavailable")


def _read_beams(beam_set, plan_data):
    machine = _text(_attr(_attr(beam_set, "MachineReference"), "MachineName"))       # [verify]
    technique = _text(_attr(beam_set, "DeliveryTechnique"))                          # [verify]
    for index, beam in enumerate(list(_attr(beam_set, "Beams") or []), start=1):     # [example]
        number = _attr(beam, "Number")                                               # [verify]
        try:
            number = int(number) if number is not None else index
        except (TypeError, ValueError):
            number = index
        mu = _attr(beam, "BeamMU")                                                   # [verify] MU per fraction
        plan_data.beams.append(BeamInfo(
            number=number, name=_text(_attr(beam, "Name")) or None,                  # [verify]
            mu=float(mu) if _finite(mu) else None, machine=machine or None, technique=technique or None))


def _read_examination(examination, plan_data):
    dt = None
    try:
        dt = _as_utc(examination.GetExaminationDateTime())                          # [verify]
    except Exception:
        dt = None
    dicom = None
    try:
        dicom = examination.GetAcquisitionDataFromDicom()                            # [verify]
    except Exception:
        dicom = None
    if dt is None:
        dt = _as_utc(_lookup(_lookup(dicom, "StudyModule"), "StudyDateTime"))        # [verify]
    plan_data.examination_datetime = dt
    equipment = _lookup(dicom, "EquipmentModule")                                    # [verify]
    imaging = _attr(_attr(_attr(examination, "EquipmentInfo"), "ImagingSystemReference"), "ImagingSystemName")  # [verify]
    plan_data.ct_device = CtDeviceInfo(
        manufacturer=_text(_lookup(equipment, "Manufacturer")) or None,
        model=_text(_lookup(equipment, "ManufacturersModelName")) or None,
        station=_text(_lookup(equipment, "StationName")) or None,
        imaging_system=_text(imaging) or None)


def _read_dose_presence(total_dose, plan_data):
    if total_dose is None:
        plan_data.has_dose = False
        return
    try:
        has_values_attribute = hasattr(total_dose, "DoseValues")
    except Exception:
        has_values_attribute = False
    if has_values_attribute:
        values = _attr(total_dose, "DoseValues")                                     # [verify]
        plan_data.has_dose = values is not None
        return
    plan_data.has_dose = None


def read_plan_data(patient, case, plan, beam_set, examination=None):
    """Convert one RayStation beam set (and its plan) into PlanData. Read-only.

    ``examination`` is used only if ``BeamSet.GetPlanningExamination()`` is unavailable (live fallback);
    the fallback is recorded as run_status partial.
    """
    data = PlanData()
    data.planning_system = "RayStation"
    data.patient_id = _text(_attr(patient, "PatientID"))                             # [verify]
    name = _text(_attr(patient, "Name"))                                             # [verify] "Last^First"
    data.patient_name_terms = [p for p in name.replace("^", " ").split(" ") if p] + ([name] if name else [])
    data.course_id = _text(_attr(case, "CaseName"))                                  # [verify]
    plan_name = _text(_attr(plan, "Name"))                                           # [verify]
    label = _text(_attr(beam_set, "DicomPlanLabel"))                                 # [verify]
    data.plan_id = "%s/%s" % (plan_name, label)
    data.plan_uid = ""  # composite pseudonym; see README (no UID attribute is assumed)
    beam_sets = _attr(plan, "BeamSets")                                              # [verify]
    try:
        data.n_beam_sets_in_plan = len(list(beam_sets)) if beam_sets is not None else None
    except Exception:
        data.n_beam_sets_in_plan = None
    review = _attr(plan, "Review")                                                   # [verify]
    data.plan_approval_status = _text(_attr(review, "ApprovalStatus"))               # [verify]
    data.planning_approval_utc = _as_utc(_attr(review, "ReviewTime"))                # [verify]

    planning_exam = None
    try:
        planning_exam = beam_set.GetPlanningExamination()                            # [verify]
    except Exception:
        planning_exam = None
    if planning_exam is None:
        planning_exam = examination
        if examination is not None:
            data.partial_reasons.append("planning_examination_fallback")
    if planning_exam is None:
        data.partial_reasons.append("planning_examination_unavailable")
    else:
        _read_rois(case, planning_exam, data)
        _read_examination(planning_exam, data)

    _read_prescription(beam_set, data)
    _read_beams(beam_set, data)
    total_dose = _attr(_attr(plan, "TreatmentCourse"), "TotalDose")                  # [example]
    _read_dose_presence(total_dose, data)
    data.dose = RayStationDose(total_dose) if total_dose is not None and data.has_dose is not False else None
    if data.has_dose is None and data.dose is not None:
        external = [r for r in data.rois if (r.roi_type or "").lower() == "external" and r.has_contours is not False]
        if len(external) == 1:
            try:
                data.has_dose = data.dose.statistic(external[0].name, "Max") > 0
            except Exception:
                data.has_dose = None
    if data.has_dose is False:
        data.dose = None
        data.partial_reasons.append("dose_unavailable")
    return data


def read_current_plan(get_current_fn=None):
    """Live review: read the currently open Patient/Case/Plan/BeamSet (and Examination as fallback). [example]"""
    get_current = get_current_fn or resolve_get_current()
    objects = {}
    for object_name in ("Patient", "Case", "Plan", "BeamSet"):
        try:
            current = get_current(object_name)
        except Exception as error:
            raise RuntimeError("Unable to obtain current RayStation %s." % object_name) from error
        if current is None:
            raise RuntimeError("A current RayStation %s is required." % object_name)
        objects[object_name] = current
    try:
        examination = get_current("Examination")
    except Exception:
        examination = None
    return read_plan_data(objects["Patient"], objects["Case"], objects["Plan"], objects["BeamSet"], examination)


def current_user_id():
    """Windows login name (not a RayStation API); pseudonymised as usr_ like the Eclipse user ID."""
    return os.environ.get("USERNAME") or os.environ.get("USER") or ""
