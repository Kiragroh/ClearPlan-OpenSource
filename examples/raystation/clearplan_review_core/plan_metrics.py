"""Plan descriptors for plan_parameter (same names/methods as the ClearPlan Eclipse implementation).

All quantities come from read-only dose queries. Unavailable -> value None plus a note, never 0.
"""

import math

PAM_UNAVAILABLE_NOTE = ("PAM not computed by the RayStation example: it needs MLC leaf geometry of the "
                        "machine model, which this read-only adapter does not read.")
APERTURE_UNAVAILABLE_NOTE = "Aperture statistics need MLC leaf geometry; not computed by the RayStation example."
DURATION_UNAVAILABLE_NOTE = "No dose-rate profile is configured in the RayStation example."
SMALL_APERTURE_THRESHOLD_CM2 = 4.0


def _finite(x):
    return x is not None and isinstance(x, (int, float)) and math.isfinite(x)


def _row(scope, entity, name, value, unit, method, note=""):
    return {"scope": scope, "entity": entity or "", "name": name, "value": value if _finite(value) else None,
            "unit": unit, "method": method, "note": "" if _finite(value) else (note or "Unavailable.")}


def mu_rows(plan):
    """total_mu and mu_per_gy (plan scope)."""
    mus = [b.mu for b in plan.beams]
    total, note = None, ""
    if not plan.beams:
        note = "No treatment beams."
    elif not all(_finite(m) and m >= 0 for m in mus):
        note = "Meterset unavailable."
    else:
        total = float(sum(mus))
    per_gy, per_note = None, ""
    dpf = plan.dose_per_fraction_gy
    if total is None:
        per_note = "Total MU unavailable."
    elif not _finite(dpf) or dpf <= 0:
        per_note = "Dose per fraction unavailable."
    else:
        per_gy = total / dpf
    return [_row("plan", "", "total_mu", total, "MU", "sum_beam_meterset", note),
            _row("plan", "", "mu_per_gy", per_gy, "MU/Gy", "total_mu_per_fraction_dose", per_note)]


def _fraction_at(dose, roi_name, dose_gy):
    value = list(dose.relative_volumes_at_doses(roi_name, [dose_gy]))[0]
    if not _finite(value) or value < -1e-6 or value > 1 + 1e-6:
        raise ValueError("relative volume outside 0..1")
    return min(1.0, max(0.0, float(value)))


def target_rows(target, external, dose, rx_gy, blocked_note=""):
    """Paddick CI, reciprocal, GI, HI, target volume and Rx for one target ROI."""
    entity = target.name if target is not None else ""
    ci = gi = hi = None
    ci_note = gi_note = hi_note = blocked_note
    tv = target.volume_cc if target is not None else None
    if not blocked_note:
        if target is None:
            ci_note = gi_note = hi_note = "Target not identified."
        elif dose is None:
            ci_note = gi_note = hi_note = "No dose available."
        elif not _finite(rx_gy) or rx_gy <= 0:
            ci_note = gi_note = hi_note = "Prescription unavailable."
        elif not _finite(tv) or tv <= 0:
            ci_note = gi_note = hi_note = "Target volume is zero or unavailable."
        else:
            ext_volume = external.volume_cc if external is not None else None
            if external is None or not _finite(ext_volume) or ext_volume <= 0:
                ci_note = gi_note = "External ROI not identified unambiguously."
            else:
                try:
                    tv_piv = _fraction_at(dose, target.name, rx_gy) * tv
                    piv = _fraction_at(dose, external.name, rx_gy) * ext_volume
                    v50 = _fraction_at(dose, external.name, 0.5 * rx_gy) * ext_volume
                    if piv > 0:
                        ci = tv_piv * tv_piv / (tv * piv)
                        gi = v50 / piv
                    else:
                        ci_note = gi_note = "Prescription isodose volume is zero."
                except Exception:
                    ci_note = gi_note = "Dose query failed."
            try:
                d2, d98 = list(dose.dose_at_relative_volumes(target.name, [0.02, 0.98]))
                if _finite(d2) and _finite(d98) and d2 >= 0 and d98 >= 0:
                    hi = (d2 - d98) / rx_gy
                else:
                    hi_note = "Dose query failed."
            except Exception:
                hi_note = "Dose query failed."
    reciprocal = 1.0 / ci if _finite(ci) and ci > 0 else None
    return [
        _row("target", entity, "paddick_ci", ci, "1", "paddick_tv_rx2_over_tv_vrx", ci_note),
        _row("target", entity, "paddick_ci_reciprocal", reciprocal, "1", "inverse_paddick", ci_note or "Paddick CI is zero."),
        _row("target", entity, "gradient_index", gi, "1", "body_v50_over_v100", gi_note),
        _row("target", entity, "homogeneity_index", hi, "1", "d2_minus_d98_over_rx", hi_note),
        _row("target", entity, "target_volume_cm3", tv if _finite(tv) and tv > 0 else None, "cm3", "structure_volume",
             "Target volume unavailable."),
        _row("target", entity, "reference_dose_gy", rx_gy, "Gy", "plan_prescription", "Prescription unavailable."),
    ]


def unavailable_plan_rows(pam_entity):
    small = "fraction_cp_area_lt_4cm2_mu_weighted"
    return [
        _row("plan", pam_entity, "pam", None, "1", "pam_unknown_weighted", PAM_UNAVAILABLE_NOTE),
        _row("plan", "", "mean_aperture_cm2", None, "cm2", "mu_weighted_mean_cp_aperture", APERTURE_UNAVAILABLE_NOTE),
        _row("plan", "", "small_aperture_fraction", None, "1", small, APERTURE_UNAVAILABLE_NOTE),
    ]


def beam_rows(beam):
    entity = "beam-%d" % beam.number if isinstance(beam.number, int) else (beam.name or "")
    mu = beam.mu if _finite(beam.mu) and beam.mu >= 0 else None
    return [
        _row("beam", entity, "total_mu", mu, "MU", "beam_meterset", "Meterset unavailable."),
        _row("beam", entity, "pam", None, "1", "pam_unknown_weighted", PAM_UNAVAILABLE_NOTE),
        _row("beam", entity, "mean_aperture_cm2", None, "cm2", "mu_weighted_mean_cp_aperture", APERTURE_UNAVAILABLE_NOTE),
        _row("beam", entity, "small_aperture_fraction", None, "1", "fraction_cp_area_lt_4cm2_mu_weighted",
             APERTURE_UNAVAILABLE_NOTE),
        _row("beam", entity, "estimated_duration_s", None, "s", "dose_rate_profile_estimate", DURATION_UNAVAILABLE_NOTE),
    ]
