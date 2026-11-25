"""DVH objectives in TG-263 / ClearPlan syntax and their read-only measurement.

Accepted forms (whitespace ignored, comma or dot decimal separator):

  D<v>%[Gy|cGy|%]     dose to the hottest v percent of the ROI          e.g. D98%[Gy], D2%[Gy]
  D<v>cc[Gy|cGy|%]    dose to the hottest v cm3 (relative volume = v / ROI volume)   e.g. D0.035cc[Gy]
  Dmax / Dmin / Dmean [Gy|cGy|%]   (also Max[Gy], Min[Gy], Mean[Gy]); unit defaults to Gy
  V<d>Gy|cGy|%[%|cc]  volume receiving at least dose d (d in % = percent of Rx)      e.g. V20Gy[%]
  CV<d>Gy|cGy|%[%|cc] complement volume = ROI volume - V(d)                           e.g. CV20Gy[cc]

Output-unit "%" for D objectives and dose-unit "%" for V objectives are relative to the prescribed
total dose (Rx). "cc" is logged as "cm3" (ClearPlan unit code).
"""

import math
import re

from . import evaluator
from .review_log import number


class ObjectiveError(ValueError):
    pass


class Objective(object):
    def __init__(self, kind, canonical, output_unit, volume=None, volume_unit=None, dose=None, dose_unit=None,
                 statistic=None, complement=False):
        self.kind = kind                  # 'dose_at_volume' | 'dose_statistic' | 'volume_at_dose'
        self.canonical = canonical
        self.output_unit = output_unit    # 'Gy' | 'cGy' | '%' | 'cc'
        self.volume = volume
        self.volume_unit = volume_unit    # '%' | 'cc'
        self.dose = dose
        self.dose_unit = dose_unit        # 'Gy' | 'cGy' | '%'
        self.statistic = statistic        # 'Max' | 'Min' | 'Average'
        self.complement = complement

    @property
    def log_unit(self):
        return "cm3" if self.output_unit == "cc" else self.output_unit

    def needs_rx(self):
        return (self.output_unit == "%" and self.kind != "volume_at_dose") or self.dose_unit == "%"


_NUM = r"(?P<n>[0-9]+(?:\.[0-9]+)?|\.[0-9]+)"
_D_AT_V = re.compile(r"^D" + _NUM + r"(?P<vu>%|cc|cm3)\[(?P<u>Gy|cGy|%)\]$", re.IGNORECASE)
_D_STAT = re.compile(r"^D?(?P<s>max|min|mean)(?:\[(?P<u>Gy|cGy|%)\])?$", re.IGNORECASE)
_V_AT_D = re.compile(r"^(?P<c>C)?V" + _NUM + r"(?P<du>Gy|cGy|%)\[(?P<u>%|cc|cm3)\]$", re.IGNORECASE)
_STAT = {"max": "Max", "min": "Min", "mean": "Average"}
_STAT_NAME = {"Max": "Dmax", "Min": "Dmin", "Average": "Dmean"}


def _unit(text):
    lowered = text.lower()
    if lowered == "gy":
        return "Gy"
    if lowered == "cgy":
        return "cGy"
    if lowered in ("cc", "cm3"):
        return "cc"
    return "%"


def _num(text):
    value = float(text)
    if not math.isfinite(value) or value < 0:
        raise ObjectiveError("Objective parameter must be a finite non-negative number.")
    return value


def parse(text):
    if text is None:
        raise ObjectiveError("Objective is empty.")
    compact = re.sub(r"\s+", "", str(text)).replace(",", ".").replace("cm\u00b3", "cc")
    if compact == "":
        raise ObjectiveError("Objective is empty.")
    m = _D_AT_V.match(compact)
    if m:
        volume, vu, out = _num(m.group("n")), _unit(m.group("vu")), _unit(m.group("u"))
        if vu == "%" and volume > 100:
            raise ObjectiveError("Relative volume above 100 %.")
        canonical = "D%s%s[%s]" % (number(volume), "cc" if vu == "cc" else "%", out)
        return Objective("dose_at_volume", canonical, out, volume=volume, volume_unit=vu)
    m = _D_STAT.match(compact)
    if m:
        stat = _STAT[m.group("s").lower()]
        out = _unit(m.group("u")) if m.group("u") else "Gy"
        return Objective("dose_statistic", "%s[%s]" % (_STAT_NAME[stat], out), out, statistic=stat)
    m = _V_AT_D.match(compact)
    if m:
        dose, du, out = _num(m.group("n")), _unit(m.group("du")), _unit(m.group("u"))
        complement = bool(m.group("c"))
        canonical = "%sV%s%s[%s]" % ("C" if complement else "", number(dose), du, "cc" if out == "cc" else "%")
        return Objective("volume_at_dose", canonical, out, dose=dose, dose_unit=du, complement=complement)
    raise ObjectiveError("Unsupported objective syntax: %r" % text)


class Measurement(object):
    def __init__(self, value=None, note=""):
        self.value = value
        self.note = note


def _finite(x):
    return x is not None and isinstance(x, (int, float)) and math.isfinite(x)


def _check_fraction(value):
    if not _finite(value) or value < -1e-6 or value > 1 + 1e-6:
        raise ValueError("relative volume outside 0..1")
    return min(1.0, max(0.0, float(value)))


def _check_dose(value):
    if not _finite(value) or value < 0:
        raise ValueError("invalid dose value")
    return float(value)


def measure(objective, roi, dose, rx_gy):
    """Measure one objective read-only. Returns Measurement(value in output unit | None, note)."""
    if dose is None:
        return Measurement(None, "No dose available.")
    volume = roi.volume_cc
    if not _finite(volume) or volume <= 0:
        return Measurement(None, "ROI volume is zero or unavailable.")
    if objective.needs_rx() and (not _finite(rx_gy) or rx_gy <= 0):
        return Measurement(None, "Prescription unavailable for a relative-dose objective.")
    try:
        if objective.kind == "dose_statistic":
            gy = _check_dose(dose.statistic(roi.name, objective.statistic))
            return Measurement(_dose_out(gy, objective.output_unit, rx_gy))
        if objective.kind == "dose_at_volume":
            if objective.volume_unit == "%":
                relative = objective.volume / 100.0
            else:
                if objective.volume > volume:
                    return Measurement(None, "Requested absolute volume exceeds the ROI volume.")
                relative = objective.volume / volume
            gy = _check_dose(list(dose.dose_at_relative_volumes(roi.name, [relative]))[0])
            return Measurement(_dose_out(gy, objective.output_unit, rx_gy))
        # volume at dose
        if objective.dose_unit == "%":
            dose_gy = objective.dose / 100.0 * rx_gy
        else:
            dose_gy = evaluator.try_convert_dose(objective.dose, objective.dose_unit, "Gy")
        fraction = _check_fraction(list(dose.relative_volumes_at_doses(roi.name, [dose_gy]))[0])
        if objective.output_unit == "%":
            value = fraction * 100.0
        else:
            value = fraction * volume
        if objective.complement:
            value = evaluator.try_complement_volume(value, volume, "%" if objective.output_unit == "%" else "cc")
            if value is None:
                return Measurement(None, "Complement volume not defined.")
        return Measurement(value)
    except Exception:  # dose engine refused (no dose, unknown ROI, ...): unavailable, never 0
        return Measurement(None, "Dose query failed.")


def _dose_out(gy, unit, rx_gy):
    if unit == "Gy":
        return gy
    if unit == "cGy":
        return gy * 100.0
    return gy / rx_gy * 100.0
