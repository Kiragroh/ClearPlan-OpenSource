"""Patient loading for retrospective mining - VERIFY LOCALLY.

UNVALIDATED EXAMPLE - NOT FOR CLINICAL USE.

Retrospective mining must open patients that are not currently loaded. The existing read-only example
never does this, so the two PatientDB calls below are assumptions that must be verified in a
non-clinical RayStation installation before use:

  get_current("PatientDB").QueryPatientInfo(Filter={"PatientID": <id>})       [verify] query only
  get_current("PatientDB").LoadPatient(PatientInfo=<info>, AllowPatientUpgrade=False)   [verify]

LoadPatient replaces the patient open in the UI and takes the patient lock, but it is not expected to
modify patient data. AllowPatientUpgrade=False is passed so that an older patient is NOT upgraded
(an upgrade would be a write). Start mining with no patient open (or with no unsaved changes);
nothing in this package calls Save, so a modified patient would never be stored by this code.
"""

ALLOWED_RAYSTATION_METHODS = frozenset(["QueryPatientInfo", "LoadPatient"])


class LoadError(Exception):
    """Carries a short technical code for run_message (never an exception message)."""

    def __init__(self, code):
        Exception.__init__(self, code)
        self.code = code


def _value(info, key):
    try:
        return info[key]
    except Exception:
        return getattr(info, key, None)


def _by_name(items, attribute, wanted):
    matches = []
    for item in list(items or []):
        try:
            value = getattr(item, attribute, None)
        except Exception:
            value = None
        if value is not None and str(value).strip() == wanted:
            matches.append(item)
    return matches


def open_plan_readonly(get_current, patient_id, case_name, plan_name, beam_set_label=""):
    """Load one patient and return (patient, case, plan, beam_set) without any write. VERIFY LOCALLY."""
    database = get_current("PatientDB")
    if database is None:
        raise LoadError("patient_db_unavailable")
    try:
        infos = list(database.QueryPatientInfo(Filter={"PatientID": patient_id}))           # [verify]
    except Exception:
        raise LoadError("patient_query_failed")
    exact = [i for i in infos if str(_value(i, "PatientID") or "").strip() == patient_id]
    if not exact:
        raise LoadError("patient_not_found")
    if len(exact) > 1:
        raise LoadError("patient_ambiguous")
    try:
        loaded = database.LoadPatient(PatientInfo=exact[0], AllowPatientUpgrade=False)     # [verify]
    except Exception:
        raise LoadError("patient_load_failed")
    patient = loaded if loaded is not None else get_current("Patient")
    if patient is None:
        raise LoadError("patient_load_failed")
    cases = _by_name(getattr(patient, "Cases", None), "CaseName", case_name)                # [verify]
    if len(cases) != 1:
        raise LoadError("case_not_found" if not cases else "case_ambiguous")
    plans = _by_name(getattr(cases[0], "TreatmentPlans", None), "Name", plan_name)          # [verify]
    if len(plans) != 1:
        raise LoadError("plan_not_found" if not plans else "plan_ambiguous")
    beam_sets = list(getattr(plans[0], "BeamSets", None) or [])                             # [verify]
    if beam_set_label:
        beam_sets = _by_name(beam_sets, "DicomPlanLabel", beam_set_label)                   # [verify]
    if len(beam_sets) != 1:
        raise LoadError("beam_set_not_found" if not beam_sets else "beam_set_ambiguous")
    return patient, cases[0], plans[0], beam_sets[0]
