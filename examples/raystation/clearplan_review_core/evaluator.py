"""Port of ClearPlan.Core.Constraints.PqmNumericEvaluator (strict, culture-independent).

Statuses: "Goal", "Variation", "Not met", "Not evaluated". Comma or dot are decimal separators,
never grouping separators. Anything that cannot be parsed strictly is "Not evaluated" and is never
counted as passed.
"""

import math
import re
from decimal import Decimal, InvalidOperation

GOAL, VARIATION, NOT_MET, NOT_EVALUATED = "Goal", "Variation", "Not met", "Not evaluated"
LOG_STATUS = {GOAL: "goal", VARIATION: "variation", NOT_MET: "not_met", NOT_EVALUATED: "not_evaluated"}
COMPARATORS = ("<=", ">=", "<", ">", "=")

_NUMBER = r"[-+]?(?:\d+(?:[.,]\d+)?|[.,]\d+)"
_ACHIEVED = re.compile(r"^\s*(?P<value>" + _NUMBER + r")\s*(?:Gy|cGy|%|cc|cm3|cm\u00b3)?\s*$")
_GOAL = re.compile(r"^\s*(?P<operator><=|>=|<|>|=)\s*(?P<value>" + _NUMBER + r")\s*$")
# decimal.TryParse(AllowLeadingSign | AllowDecimalPoint, InvariantCulture) after the comma rule
_DECIMAL = re.compile(r"^[+-]?(?:[0-9]+\.?[0-9]*|\.[0-9]+)$")


def parse_nullable_decimal(raw):
    """Port of ConstraintValueNormalizer.ParseNullableDecimal: None for blank, ValueError if invalid."""
    if raw is None:
        return None
    value = str(raw)
    if value.strip() == "":
        return None
    value = value.strip()
    if "," in value and "." not in value:
        value = value.replace(",", ".")
    if not _DECIMAL.match(value):
        raise ValueError("Constraint value '%s' is not a valid invariant decimal." % value)
    try:
        return Decimal(value)
    except InvalidOperation:
        raise ValueError("Constraint value '%s' is not a valid invariant decimal." % value)


def try_parse_number(text):
    """Return float or None (TryParseNumber)."""
    try:
        parsed = parse_nullable_decimal(text)
    except ValueError:
        return None
    return None if parsed is None else float(parsed)


def try_parse_achieved(text):
    match = _ACHIEVED.match(text or "")
    if not match:
        return None
    value = try_parse_number(match.group("value"))
    return value if value is not None and value >= 0 else None


def try_parse_goal(text):
    """Return (comparator, value) or None."""
    match = _GOAL.match(text or "")
    if not match:
        return None
    value = try_parse_number(match.group("value"))
    if value is None or value < 0:
        return None
    return match.group("operator"), value


def _matches(actual, limit, comparator):
    if comparator == "<":
        return actual < limit
    if comparator == "<=":
        return actual <= limit
    if comparator == ">":
        return actual > limit
    if comparator == ">=":
        return actual >= limit
    if comparator == "=":
        return actual == limit
    return False


def evaluate(achieved, goal, variation):
    has_variation = variation is not None and str(variation).strip() != ""
    actual = try_parse_achieved(achieved)
    parsed_goal = try_parse_goal(goal)
    allowed = try_parse_number(variation) if has_variation else 0.0
    if actual is None or parsed_goal is None or (has_variation and (allowed is None or allowed < 0)):
        return NOT_EVALUATED
    comparator, limit = parsed_goal
    if _matches(actual, limit, comparator):
        return GOAL
    return VARIATION if has_variation and _matches(actual, allowed, comparator) else NOT_MET


def evaluate_for_confirmed_scope(achieved, goal, variation, scope_confirmed):
    return evaluate(achieved, goal, variation) if scope_confirmed else NOT_EVALUATED


def try_get_goal_ratio(achieved, goal):
    actual = try_parse_achieved(achieved)
    parsed = try_parse_goal(goal)
    if actual is None or parsed is None or parsed[1] == 0:
        return None
    ratio = actual / parsed[1] * 100
    return None if math.isnan(ratio) or math.isinf(ratio) else ratio


def try_convert_dose(value, source_unit, target_unit):
    if value is None or math.isnan(value) or math.isinf(value) or value < 0:
        return None
    if source_unit == "%" and target_unit == "%":
        return value
    if source_unit not in ("Gy", "cGy") or target_unit not in ("Gy", "cGy"):
        return None
    converted = value * (0.01 if source_unit == "cGy" else 1) * (100 if target_unit == "cGy" else 1)
    return None if math.isnan(converted) or math.isinf(converted) else converted


def _finite(x):
    return x is not None and not math.isnan(x) and not math.isinf(x)


def try_complement_volume(volume_at_dose, organ_volume_cc, volume_unit, minimum_volume=None):
    """Complement in its own unit: 100 - V% or organ volume - Vcc (TryComplementVolume)."""
    if not _finite(organ_volume_cc) or organ_volume_cc <= 0 or volume_unit not in ("%", "cc"):
        return None
    total = 100.0 if volume_unit == "%" else organ_volume_cc
    if not _finite(volume_at_dose) or volume_at_dose < 0 or volume_at_dose > total:
        return None
    if minimum_volume is not None and (not _finite(minimum_volume) or minimum_volume < 0 or minimum_volume > total):
        return None
    return total - volume_at_dose


def try_get_complement_goal_ratio(achieved, goal, organ_volume_cc, volume_unit):
    actual = try_parse_achieved(achieved)
    parsed = try_parse_goal(goal)
    if actual is None or parsed is None:
        return None
    actual_c = try_complement_volume(actual, organ_volume_cc, volume_unit)
    limit_c = try_complement_volume(parsed[1], organ_volume_cc, volume_unit)
    if actual_c is None or limit_c is None or limit_c == 0:
        return None
    ratio = actual_c / limit_c * 100
    return None if math.isnan(ratio) or math.isinf(ratio) else ratio


def format_invariant(value):
    """Invariant fixed-point text (shortest round-trip digits, never exponent notation)."""
    if not _finite(value):
        return ""
    text = format(Decimal(repr(float(value))), "f")
    return "0" if text in ("-0", "-0.0") else text


def evaluate_numeric(achieved_value, comparator, goal_value, variation_value, scope_confirmed=True):
    """Numeric front-end: builds the same strings the C# evaluator receives."""
    achieved = format_invariant(achieved_value) if achieved_value is not None else ""
    goal = "" if goal_value is None or comparator not in COMPARATORS else comparator + format_invariant(goal_value)
    variation = "" if variation_value is None else format_invariant(variation_value)
    return evaluate_for_confirmed_scope(achieved, goal, variation, scope_confirmed)
