"""TPS-neutral plan description filled by an adapter (RayStation, tests, ...).

Units: dose in Gy, volume in cm3, relative volume as fraction 0..1, times as timezone-aware datetimes.
``None`` always means "not available" and is never replaced by 0.
"""


class RoiInfo(object):
    def __init__(self, name, roi_type=None, has_contours=None, volume_cc=None):
        self.name = name
        self.roi_type = roi_type
        self.has_contours = has_contours
        self.volume_cc = volume_cc


class BeamInfo(object):
    def __init__(self, number=None, name=None, mu=None, machine=None, technique=None):
        self.number = number
        self.name = name
        self.mu = mu
        self.machine = machine
        self.technique = technique


class CtDeviceInfo(object):
    def __init__(self, manufacturer=None, model=None, station=None, imaging_system=None):
        self.manufacturer = manufacturer
        self.model = model
        self.station = station
        self.imaging_system = imaging_system

    def is_empty(self):
        return not any([self.manufacturer, self.model, self.station, self.imaging_system])


class DoseQuery(object):
    """Read-only dose interface (Gy, fractions). Implementations raise on failure."""

    def dose_at_relative_volumes(self, roi_name, relative_volumes):
        raise NotImplementedError

    def relative_volumes_at_doses(self, roi_name, doses_gy):
        raise NotImplementedError

    def statistic(self, roi_name, kind):
        """kind in {'Max', 'Min', 'Average'}."""
        raise NotImplementedError


class PlanData(object):
    """Everything the core needs about one reviewed plan (beam set)."""

    def __init__(self):
        self.planning_system = ""
        # raw identifiers (memory only; pseudonymised by the log writer)
        self.patient_id = ""
        self.patient_name_terms = []
        self.course_id = ""          # RayStation: case name
        self.plan_id = ""            # RayStation: "<plan name>/<beam set label>"
        self.plan_uid = ""           # empty -> composite pseudonym (see README)
        self.total_dose_gy = None    # prescribed total dose of the reviewed beam set
        self.fractions = None
        self.n_beam_sets_in_plan = None
        self.prescription_target = None
        self.rois = []
        self.beams = []
        self.examination_datetime = None
        self.ct_device = CtDeviceInfo()
        self.plan_approval_status = ""
        self.planning_approval_utc = None
        self.plan_created_utc = None
        self.has_dose = None         # True / False / None (unknown)
        self.dose = None             # DoseQuery or None
        self.partial_reasons = []    # short technical codes; make run_status 'partial'

    @property
    def dose_per_fraction_gy(self):
        if self.total_dose_gy is None or not self.fractions:
            return None
        return self.total_dose_gy / self.fractions

    def roi(self, name):
        for roi in self.rois:
            if roi.name == name:
                return roi
        return None
