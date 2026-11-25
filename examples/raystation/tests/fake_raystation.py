"""Fake RayStation object model for tests outside RayStation (simulated API only).

Attribute and method names mirror what raystation_adapter.py / raystation_patient_loader.py read.
Unknown attributes raise AttributeError; every method call is recorded in CALLS so that tests can
assert that only allowlisted read-only methods were used. Doses are handled in cGy like RayStation.
"""

import types
from datetime import datetime, timezone

CALLS = []


def record(name):
    CALLS.append(name)


class Strict(object):
    def __getattr__(self, name):
        raise AttributeError(name)


class Dvh(object):
    """Piecewise-linear cumulative DVH: [(dose_cGy, relative_volume)], volume non-increasing."""

    def __init__(self, points):
        self.points = list(points)

    def rel_at(self, dose):
        pts = self.points
        if dose <= pts[0][0]:
            return pts[0][1]
        for (d0, v0), (d1, v1) in zip(pts, pts[1:]):
            if d0 <= dose <= d1:
                return v0 if d1 == d0 else v0 + (v1 - v0) * (dose - d0) / (d1 - d0)
        return 0.0

    def dose_at(self, rel):
        pts = self.points
        if rel >= 1.0:
            return max(d for d, v in pts if v >= 1.0)
        for (d0, v0), (d1, v1) in zip(pts, pts[1:]):
            if v0 >= rel >= v1 and v0 != v1:
                return d0 + (v0 - rel) / (v0 - v1) * (d1 - d0)
        return pts[-1][0]

    def mean(self):
        return sum((v0 + v1) / 2.0 * (d1 - d0) for (d0, v0), (d1, v1) in zip(self.points, self.points[1:]))


class FakeTotalDose(Strict):
    def __init__(self, dvhs, has_values=True):
        self._dvhs = dvhs
        self.DoseValues = object() if has_values else None

    def _dvh(self, name):
        if self.DoseValues is None:
            raise RuntimeError("no dose")
        return self._dvhs[name]  # KeyError for unknown ROI, like a refused query

    def GetDoseStatistic(self, RoiName, DoseType):
        record("GetDoseStatistic")
        dvh = self._dvh(RoiName)
        if DoseType == "Max":
            return dvh.points[-1][0]
        if DoseType == "Min":
            return dvh.dose_at(1.0)
        if DoseType == "Average":
            return dvh.mean()
        raise ValueError(DoseType)

    def GetRelativeVolumeAtDoseValues(self, RoiName, DoseValues):
        record("GetRelativeVolumeAtDoseValues")
        dvh = self._dvh(RoiName)
        return [dvh.rel_at(float(d)) for d in DoseValues]

    def GetDoseAtRelativeVolumes(self, RoiName, RelativeVolumes):
        record("GetDoseAtRelativeVolumes")
        dvh = self._dvh(RoiName)
        return [dvh.dose_at(float(v)) for v in RelativeVolumes]


class FakeRoi(Strict):
    def __init__(self, name, roi_type):
        self.Name = name
        self.Type = roi_type


class FakeGeometry(Strict):
    def __init__(self, volume, has_contours=True):
        self._volume = volume
        self._has = has_contours

    def HasContours(self):
        record("HasContours")
        return self._has

    def GetRoiVolume(self):
        record("GetRoiVolume")
        if not self._has:
            raise RuntimeError("no contours")
        return self._volume


class FakeExamination(Strict):
    def __init__(self, name, when, dicom, imaging_system):
        self.Name = name
        self._when = when
        self._dicom = dicom
        self.EquipmentInfo = types.SimpleNamespace(
            ImagingSystemReference=types.SimpleNamespace(ImagingSystemName=imaging_system))

    def GetExaminationDateTime(self):
        record("GetExaminationDateTime")
        return self._when

    def GetAcquisitionDataFromDicom(self):
        record("GetAcquisitionDataFromDicom")
        return self._dicom


class FakeBeam(Strict):
    def __init__(self, number, name, mu):
        self.Number = number
        self.Name = name
        self.BeamMU = mu


class FakeBeamSet(Strict):
    def __init__(self, label, fractions, rx_cgy, target, beams, exam, machine="ExampleLinac", technique="SMLC"):
        self.DicomPlanLabel = label
        self.FractionationPattern = types.SimpleNamespace(NumberOfFractions=fractions)
        self.Prescription = types.SimpleNamespace(PrimaryDosePrescription=types.SimpleNamespace(
            DoseValue=rx_cgy, OnStructure=types.SimpleNamespace(Name=target)))
        self.Beams = beams
        self.MachineReference = types.SimpleNamespace(MachineName=machine)
        self.DeliveryTechnique = technique
        self._exam = exam

    def GetPlanningExamination(self):
        record("GetPlanningExamination")
        return self._exam


class FakePlan(Strict):
    def __init__(self, name, total_dose, beam_sets, approval="Approved", review_time=None):
        self.Name = name
        self.TreatmentCourse = types.SimpleNamespace(TotalDose=total_dose)
        self.BeamSets = beam_sets
        self.Review = types.SimpleNamespace(ApprovalStatus=approval, ReviewTime=review_time)


class FakeCase(Strict):
    def __init__(self, name, rois, geometries, exam_name, plans):
        self.CaseName = name
        self.PatientModel = types.SimpleNamespace(
            RegionsOfInterest=rois,
            StructureSets={exam_name: types.SimpleNamespace(RoiGeometries=geometries)})
        self.TreatmentPlans = plans


class FakePatient(Strict):
    def __init__(self, patient_id, name, cases):
        self.PatientID = patient_id
        self.Name = name
        self.Cases = cases


class FakePatientDB(Strict):
    def __init__(self, patients, state):
        self._patients = patients
        self._state = state
        self.load_calls = []

    def QueryPatientInfo(self, Filter):
        record("QueryPatientInfo")
        pid = Filter["PatientID"]
        return [{"PatientID": p.PatientID, "LastName": "x"} for p in self._patients if p.PatientID == pid]

    def LoadPatient(self, PatientInfo, AllowPatientUpgrade):
        record("LoadPatient")
        self.load_calls.append((PatientInfo["PatientID"], AllowPatientUpgrade))
        patient = [p for p in self._patients if p.PatientID == PatientInfo["PatientID"]][0]
        self._state["Patient"] = patient
        return None


PTV = [(0, 1.0), (2400, 1.0), (2500, 0.95), (2600, 0.10), (2700, 0.0)]
RECTUM = [(0, 1.0), (500, 0.6), (1800, 0.2), (2500, 0.02), (2600, 0.0)]
EXTERNAL = [(0, 1.0), (100, 0.5), (1250, 0.012), (2500, 0.0065), (2700, 0.0)]
EXAM_TIME = datetime(2026, 10, 20, 9, 0, tzinfo=timezone.utc)
REVIEW_TIME = datetime(2026, 10, 27, 10, 0, tzinfo=timezone.utc)
DICOM = {"EquipmentModule": {"Manufacturer": "ExampleVendor", "ManufacturersModelName": "ExampleCT 1",
                             "StationName": "CT-EXAMPLE-01"},
         "StudyModule": {"StudyDateTime": EXAM_TIME}}


def build_phantom(patient_id="RS-PHANTOM-001", patient_name="PHANTOM^RAYSTATION", case_name="Case 1",
                  plan_name="1AA_PCA_CP", label="1AA_PCA_CP", fractions=5, rx_cgy=2500.0,
                  extra_rois=(), rectum_volume=50.0, exam_time=EXAM_TIME, dicom=DICOM, has_dose=True,
                  beam_names=("1 AP", "2 PA"), n_beam_sets=1, imaging_system="CT-ExampleCT1-120kV",
                  external_type="External"):
    """Phantom-like prostate plan: 5 x 5 Gy, PTV 100 cm3, rectum 50 cm3, external 20 000 cm3, two beams."""
    rois = [FakeRoi("PTV", "Ptv"), FakeRoi("Rectum", "Organ"), FakeRoi("External", external_type)]
    geometries = {"PTV": FakeGeometry(100.0), "Rectum": FakeGeometry(rectum_volume),
                  "External": FakeGeometry(20000.0)}
    dvhs = {"PTV": Dvh(PTV), "Rectum": Dvh(RECTUM), "External": Dvh(EXTERNAL)}
    for name, roi_type, volume, has_contours, points in extra_rois:
        rois.append(FakeRoi(name, roi_type))
        geometries[name] = FakeGeometry(volume, has_contours)
        dvhs[name] = Dvh(points)
    exam = FakeExamination("CT 1", exam_time, dicom, imaging_system)
    beams = [FakeBeam(1, beam_names[0], 120.0), FakeBeam(2, beam_names[1], 118.0)]
    beam_set = FakeBeamSet(label, fractions, rx_cgy, "PTV", beams, exam)
    beam_sets = [beam_set] + [FakeBeamSet(label + "_b%d" % i, fractions, rx_cgy, "PTV", beams, exam)
                              for i in range(2, n_beam_sets + 1)]
    plan = FakePlan(plan_name, FakeTotalDose(dvhs, has_dose), beam_sets,
                    review_time=datetime(2026, 10, 26, 15, 0, tzinfo=timezone.utc))
    case = FakeCase(case_name, rois, geometries, exam.Name, [plan])
    patient = FakePatient(patient_id, patient_name, [case])
    objects = {"Patient": patient, "Case": case, "Plan": plan, "BeamSet": beam_set, "Examination": exam}

    def get_current(name):
        record("get_current:" + name)
        return objects[name]

    return objects, get_current


def build_database(patients):
    state = {}
    database = FakePatientDB(patients, state)

    def get_current(name):
        record("get_current:" + name)
        if name == "PatientDB":
            return database
        return state.get(name)

    return database, get_current
