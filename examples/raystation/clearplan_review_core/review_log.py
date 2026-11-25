"""Review-log contract (schema version 1) - Python port of ClearPlan.Core/ReviewLog.

Byte-compatible with the C# writer (ReviewLogSchema/Csv/Pseudonymizer/RecordBuilder/
Writer): identical headers and column order, UTF-8 without BOM, RFC 4180 with CRLF,
invariant numbers in .NET "R" form, ISO-8601 UTC with milliseconds and literal Z,
empty field = unavailable (never 0), HMAC-SHA256 pseudonyms with the same 32-byte
base64 key file and prefixes pat_/crs_/pln_/usr_. One key file and one log
database can therefore be shared by Eclipse (ESAPI) and RayStation reviews.
"""

import base64
import errno
import hashlib
import hmac
import math
import os
import random
import re
import time
import uuid
from datetime import datetime, timezone

SCHEMA_VERSION = 1

REVIEW = "review"
GOAL_RESULT = "goal_result"
PLAN_PARAMETER = "plan_parameter"
CHECK_FINDING = "check_finding"
RULESET_REVISION = "ruleset_revision"

COLUMNS = {
    REVIEW: (
        "schema_version", "review_id", "mode", "started_utc", "finished_utc", "clearplan_version", "build_id",
        "identifier_mode", "user_key", "patient_key", "course_key", "plan_key", "plan_label", "planning_item_type",
        "plan_created_utc", "plan_approval_status", "planning_approval_utc", "treatment_approval_utc",
        "machine_id", "technique", "mlc_model", "total_dose_gy", "dose_per_fraction_gy", "fractions",
        "goal_table_id", "goal_table_sha256", "ruleset_sha256",
        "n_goals", "n_goal", "n_variation", "n_not_met", "n_not_evaluated",
        "n_checks", "n_checks_excluded", "n_findings_open",
        "report_written", "report_sha256", "report_filed", "run_status", "run_message",
    ),
    GOAL_RESULT: (
        "review_id", "row_no", "stable_id", "source_kind", "template_code", "template_structure",
        "resolved_structure", "mapping_status", "objective", "comparator", "goal_value", "variation_value",
        "achieved_value", "unit", "status", "priority", "source_label",
    ),
    PLAN_PARAMETER: (
        "review_id", "row_no", "scope", "entity", "name", "value", "unit", "method", "availability", "note",
    ),
    CHECK_FINDING: (
        "review_id", "row_no", "check_code", "category", "severity", "status", "included",
        "observed_value", "expected_value", "observed_numeric", "expected_numeric", "unit", "message",
    ),
    RULESET_REVISION: (
        "review_id", "config_key", "category", "revision_number", "sha256", "revision_created_utc", "availability",
    ),
}

# Child tables first, the review row last (a review row implies complete children).
WRITE_ORDER = (RULESET_REVISION, GOAL_RESULT, PLAN_PARAMETER, CHECK_FINDING, REVIEW)

MODES = ("live", "headless", "retrospective")
RUN_OK, RUN_PARTIAL, RUN_FAILED = "ok", "partial", "failed"
GOAL_STATUSES = ("goal", "variation", "not_met", "not_evaluated")
CHECK_STATUSES = ("ok", "variation", "fail", "info", "not_evaluated")
OPEN_CHECK_STATUSES = ("variation", "fail", "not_evaluated")
SEVERITIES = ("none", "info", "warning", "error")
MAPPING_STATUSES = ("resolved", "ambiguous", "unresolved")
SOURCE_KINDS = ("constraint_table", "eclipse_clinical_goal", "default_rule")
ECLIPSE_GOAL_SOURCE_PREFIX = "Eclipse Clinical Goals"
DEFAULT_RULE_SOURCE_LABEL = None  # the ClearPlan default target rule is not part of this example

LOCK_FILE_NAME = ".review-log.lock"
MINIMUM_KEY_BYTES = 32
PSEUDONYM_HEX_LENGTH = 32
SCRUB_MAXIMUM_LENGTH = 2000
REDACTED = "[redacted]"


# --------------------------------------------------------------------------------------------
# CSV formatting (ReviewLogCsv)
# --------------------------------------------------------------------------------------------

def escape(value):
    """RFC 4180 field: quote if it contains , " CR LF or starts/ends with whitespace."""
    if value is None or value == "":
        return ""
    value = str(value)
    quote = any(c in value for c in ',"\r\n') or value[0].isspace() or value[-1].isspace()
    return '"' + value.replace('"', '""') + '"' if quote else value


def format_line(fields):
    return ",".join(escape(f) for f in fields) + "\r\n"


def number(value):
    """Finite number in .NET Framework ``double.ToString("R", Invariant)`` form; else empty.

    .NET Framework "R": 15 significant digits if they round-trip, otherwise 17; trailing zeros
    removed; scientific notation as ``1E-05`` / ``1E+16``. Negative zero is written as ``0``.
    """
    if value is None:
        return ""
    try:
        value = float(value)
    except (TypeError, ValueError):
        return ""
    if math.isnan(value) or math.isinf(value):
        return ""
    if value == 0.0:
        return "0"
    text = "%.15G" % value
    if float(text) != value:
        text = "%.17G" % value
    return text


def integer(value):
    return "" if value is None else str(int(value))


def flag(value):
    return "" if value is None else ("1" if value else "0")


def utc(value):
    """ISO 8601 UTC with millisecond precision (truncated like .NET 'fff') and literal Z."""
    if value is None:
        return ""
    if not isinstance(value, datetime):
        raise TypeError("utc() expects a datetime")
    if value.tzinfo is None or value.utcoffset() is None:
        raise ValueError("utc() requires a timezone-aware datetime")
    value = value.astimezone(timezone.utc)
    return value.strftime("%Y-%m-%dT%H:%M:%S.") + "%03dZ" % (value.microsecond // 1000)


_INVARIANT_FLOAT = re.compile(r"^\s*[+-]?(?:[0-9]+\.?[0-9]*|\.[0-9]+)(?:[eE][+-]?[0-9]+)?\s*$")


def try_parse_invariant_number(text):
    """Port of ReviewLogCsv.TryParseInvariantNumber (NumberStyles.Float, invariant culture)."""
    if text is None or str(text).strip() == "":
        return None
    if not _INVARIANT_FLOAT.match(str(text)):
        return None
    value = float(str(text).strip())
    return None if math.isnan(value) or math.isinf(value) else value


def parse_csv(text):
    """Minimal RFC 4180 reader (port of ReviewLogCsv.Parse); blank lines are skipped."""
    rows = []
    if not text:
        return rows
    if text[0] == "\ufeff":
        text = text[1:]
    row, field, in_quotes, started = [], [], False, False
    i, n = 0, len(text)
    while i < n:
        c = text[i]
        if in_quotes:
            if c == '"':
                if i + 1 < n and text[i + 1] == '"':
                    field.append('"')
                    i += 1
                else:
                    in_quotes = False
            else:
                field.append(c)
        elif c == '"' and not field:
            in_quotes, started = True, True
        elif c == ",":
            row.append("".join(field))
            field, started = [], False
        elif c in "\r\n":
            if c == "\r" and i + 1 < n and text[i + 1] == "\n":
                i += 1
            row.append("".join(field))
            field, started = [], False
            if not (len(row) == 1 and row[0] == ""):
                rows.append(row)
            row = []
        else:
            field.append(c)
            started = True
        i += 1
    if in_quotes:
        raise ValueError("Unterminated quoted CSV field.")
    if started or field or row:
        row.append("".join(field))
        if not (len(row) == 1 and row[0] == ""):
            rows.append(row)
    return rows


def sha256_hex(data):
    if isinstance(data, str):
        data = data.encode("utf-8")
    return hashlib.sha256(data or b"").hexdigest()


def file_sha256(path):
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        for block in iter(lambda: handle.read(1 << 20), b""):
            digest.update(block)
    return digest.hexdigest()


# --------------------------------------------------------------------------------------------
# Pseudonymisation (ReviewLogPseudonymizer, ReviewLogTextScrubber)
# --------------------------------------------------------------------------------------------

def _upper_invariant(text):
    # .NET ToUpperInvariant uses simple (1:1) case mapping; Python's str.upper may expand (e.g. sharp s).
    out = []
    for ch in text:
        up = ch.upper()
        out.append(up if len(up) == 1 else ch)
    return "".join(out)


def _normalize(value, case_insensitive):
    trimmed = (value or "").strip()
    return _upper_invariant(trimmed) if case_insensitive else trimmed


def _strip_domain(value):
    if value is None or value.strip() == "":
        return ""
    parts = [p for p in value.split("\\") if p]
    return value if not parts else parts[-1]


class Pseudonymizer(object):
    """Keyed pseudonyms (HMAC-SHA256, truncated to 128 bit). Same key file as the C# writer."""

    def __init__(self, key, log_clear_identifiers=False):
        if not log_clear_identifiers and (key is None or len(key) < MINIMUM_KEY_BYTES):
            raise ValueError("A review-log key of at least 32 bytes is required for pseudonymisation.")
        self._key = None if key is None else bytes(key)
        self._clear = bool(log_clear_identifiers)

    @property
    def clear_identifiers(self):
        return self._clear

    @property
    def identifier_mode(self):
        return "clear" if self._clear else "hmac_sha256"

    def patient(self, patient_id):
        return self._key_for("pat", _normalize(patient_id, True), patient_id)

    def course(self, patient_id, course_id):
        if course_id is None or course_id.strip() == "":
            return ""
        return self._key_for("crs", _normalize(patient_id, True) + "\u001f" + _normalize(course_id, False), course_id)

    def plan(self, plan_uid, patient_id, course_id, plan_id):
        if plan_uid is not None and plan_uid.strip() != "":
            return self._key_for("pln", "uid\u001f" + plan_uid.strip(), plan_uid)
        if plan_id is None or plan_id.strip() == "":
            return ""
        composite = "id\u001f" + _normalize(patient_id, True) + "\u001f" + _normalize(course_id, False) + \
            "\u001f" + _normalize(plan_id, False)
        return self._key_for("pln", composite, (course_id or "") + "/" + plan_id)

    def user(self, user_id):
        return self._key_for("usr", _normalize(_strip_domain(user_id), True), user_id)

    def _key_for(self, prefix, normalized, clear_value):
        if clear_value is None or clear_value.strip() == "" or normalized == "":
            return ""
        if self._clear:
            return clear_value.strip()
        mac = hmac.new(self._key, (prefix + "\u001e" + normalized).encode("utf-8"), hashlib.sha256).digest()
        return prefix + "_" + mac[:PSEUDONYM_HEX_LENGTH // 2].hex()


def load_or_create_key(key_file_path):
    """Read the base64 key file or create a new random 32-byte key exactly once (race safe)."""
    if not key_file_path or str(key_file_path).strip() == "":
        raise ValueError("A review-log key file path is required.")
    path = os.fspath(key_file_path)
    for _ in range(20):
        if os.path.exists(path):
            return _read_key(path)
        directory = os.path.dirname(os.path.abspath(path))
        if directory:
            os.makedirs(directory, exist_ok=True)
        key = os.urandom(MINIMUM_KEY_BYTES)
        try:
            fd = os.open(path, os.O_WRONLY | os.O_CREAT | os.O_EXCL | getattr(os, "O_BINARY", 0))
        except OSError as error:
            if error.errno != errno.EEXIST:
                raise
            time.sleep(0.025)
            continue
        try:
            os.write(fd, base64.b64encode(key))
            os.fsync(fd)
        finally:
            os.close(fd)
        return key
    return _read_key(path)


def _read_key(path):
    last = None
    for _ in range(21):
        try:
            with open(path, "rb") as handle:
                text = handle.read().decode("ascii").strip()
            key = base64.b64decode(text, validate=True)
            if len(key) < MINIMUM_KEY_BYTES:
                raise ValueError("Review-log key file is shorter than 32 bytes.")
            return key
        except (OSError, ValueError, UnicodeDecodeError) as error:
            if isinstance(error, ValueError) and "shorter" in str(error):
                raise
            last = error
            time.sleep(0.025)
    raise last


class TextScrubber(object):
    """Replaces known identifying terms in free text with [redacted]; truncates to 2000 characters."""

    def __init__(self, sensitive_terms):
        terms = []
        for term in sensitive_terms or ():
            if term is None or str(term).strip() == "":
                continue
            trimmed = str(term).strip()
            if len(trimmed) >= 3 and trimmed not in terms:
                terms.append(trimmed)
        terms.sort(key=len, reverse=True)
        self._terms = [(t, _upper_invariant(t)) for t in terms]

    def scrub(self, text):
        if not text:
            return ""
        text = str(text)
        upper = _upper_invariant(text)
        out, position = [], 0
        while position < len(text):
            match = None
            for term, term_upper in self._terms:
                if len(term) <= len(text) - position and upper[position:position + len(term)] == term_upper:
                    match = term
                    break
            if match is None:
                out.append(text[position])
                position += 1
            else:
                out.append(REDACTED)
                position += len(match)
        result = "".join(out)
        return result[:SCRUB_MAXIMUM_LENGTH]


# --------------------------------------------------------------------------------------------
# Rule-set provenance (ReviewLogRulesetResolver)
# --------------------------------------------------------------------------------------------

MAXIMUM_HASHED_BYTES = 64 * 1024 * 1024


class ConfigurationInput(object):
    def __init__(self, config_key, category, sha256=None, availability=None, revision_number=None,
                 revision_created_utc=None):
        self.config_key = config_key
        self.category = category
        self.sha256 = sha256
        self.availability = availability
        self.revision_number = revision_number
        self.revision_created_utc = revision_created_utc


def describe_config_file(config_key, category, path):
    """SHA-256 of the configuration file actually used (managed revision metadata is not resolved here)."""
    item = ConfigurationInput(config_key, category)
    if path is None or str(path).strip() == "":
        item.availability = "not_configured"
        return item
    try:
        if os.path.isdir(path):
            item.availability = "not_hashed"
            return item
        if not os.path.isfile(path):
            item.availability = "unavailable"
            return item
        if os.path.getsize(path) > MAXIMUM_HASHED_BYTES:
            item.availability = "not_hashed"
            return item
        item.sha256 = file_sha256(path)
        item.availability = "available"
    except OSError:
        item.availability = "unavailable"
    return item


def ruleset_sha256(inputs, goal_table_sha256):
    """Combined fingerprint: sorted ``config_key=sha256`` lines plus ``goal_table=<sha>`` (C# RulesetSha256)."""
    lines = []
    for item in inputs or ():
        if item is None:
            continue
        value = item.sha256 if item.sha256 is not None else (item.availability if item.availability is not None else "")
        lines.append((item.config_key or "") + "=" + value.lower())
    lines.sort()
    if goal_table_sha256:
        lines.append("goal_table=" + goal_table_sha256)
    return "" if not lines else sha256_hex("\n".join(lines))


def goal_table_sha256(goal_rows):
    """SHA-256 over canonical definitions of the constraint-table goal rows evaluated in this review."""
    lines = []
    for row in goal_rows:
        if source_kind(row.get("source_label")) != "constraint_table":
            continue
        lines.append("\u001f".join([
            row.get("template_code") or "", row.get("template_structure") or "", row.get("objective") or "",
            row.get("comparator") or "", number(row.get("goal_value")), number(row.get("variation_value")),
            row.get("unit") or "",
        ]))
    return "" if not lines else sha256_hex("\n".join(lines))


def source_kind(source_label):
    if source_label is None or str(source_label).strip() == "":
        return "constraint_table"
    if str(source_label).lower().startswith(ECLIPSE_GOAL_SOURCE_PREFIX.lower()):
        return "eclipse_clinical_goal"
    return "constraint_table"


# --------------------------------------------------------------------------------------------
# Record builder (ReviewLogRecordBuilder)
# --------------------------------------------------------------------------------------------

class ReviewContext(object):
    """Per-review facts that are not part of the evaluated content (identity, timing, provenance).

    Raw identifiers stay in memory; they are pseudonymised (or written clear on explicit policy).
    """

    def __init__(self, mode, review_id=None):
        if mode not in MODES:
            raise ValueError("Unknown review-log mode '%s'." % mode)
        self.mode = mode
        self.review_id = review_id or str(uuid.uuid4())
        self.started_utc = None
        self.finished_utc = None
        self.clearplan_version = ""
        self.build_id = ""
        self.user_id = ""
        self.patient_id = ""
        self.course_id = ""
        self.plan_uid = ""
        self.plan_id = ""
        self.planning_item_type = "plan"
        self.plan_created_utc = None
        self.plan_approval_status = ""
        self.planning_approval_utc = None
        self.treatment_approval_utc = None
        self.goal_table_id = ""
        self.report_written = False
        self.report_sha256 = ""
        self.report_filed = False
        self.run_status = RUN_OK
        self.run_message = ""
        self.configuration_inputs = []
        self.sensitive_terms = []
        self.n_checks_excluded = 0


class ReviewLogOptions(object):
    def __init__(self, log_clear_identifiers=False, log_plan_label=False, log_messages=True,
                 lock_timeout_seconds=5.0):
        self.log_clear_identifiers = log_clear_identifiers
        self.log_plan_label = log_plan_label
        self.log_messages = log_messages
        self.lock_timeout_seconds = lock_timeout_seconds


class ReviewContent(object):
    """Evaluated content of one review (``None`` content = failed run, like a null C# snapshot).

    goal_rows: dicts with template_code, template_structure, resolved_structure, mapping_status, objective,
      comparator, goal_value, variation_value, achieved_value, unit, status, stable_id, source_label.
    parameter_rows: dicts with scope, entity, name, value, unit, method, note.
    check_rows: dicts with check_code, category, severity, status, observed_value, expected_value, unit, message.
    """

    def __init__(self, goal_rows=None, parameter_rows=None, check_rows=None, machine_ids=None, techniques=None,
                 mlc_models=None, total_dose_gy=None, dose_per_fraction_gy=None, fractions=None):
        self.goal_rows = list(goal_rows or [])
        self.parameter_rows = list(parameter_rows or [])
        self.check_rows = list(check_rows or [])
        self.machine_ids = list(machine_ids or [])
        self.techniques = list(techniques or [])
        self.mlc_models = list(mlc_models or [])
        self.total_dose_gy = total_dose_gy
        self.dose_per_fraction_gy = dose_per_fraction_gy
        self.fractions = fractions


def _distinct(values):
    seen = sorted(set(v.strip() for v in values if v is not None and str(v).strip() != ""))
    return "|".join(seen)


def build_record(content, context, pseudonymizer, options=None):
    """Project evaluated content + context into the five tables (dict table -> list of field lists)."""
    options = options or ReviewLogOptions()
    scrubber = TextScrubber(context.sensitive_terms)

    def text(value):
        return scrubber.scrub(value) if options.log_messages else ""

    rid = context.review_id
    rows = dict((table, []) for table in WRITE_ORDER)

    for item in sorted([i for i in context.configuration_inputs if i is not None], key=lambda i: i.config_key or ""):
        rows[RULESET_REVISION].append([
            rid, item.config_key or "", item.category or "", integer(item.revision_number),
            (item.sha256 or "").lower(), utc(item.revision_created_utc), item.availability or "",
        ])

    statuses = []
    goal_sha = ""
    checks = open_findings = 0
    if content is not None:
        for row_no, g in enumerate(content.goal_rows, start=1):
            status = g.get("status") if g.get("status") in GOAL_STATUSES else "not_evaluated"
            statuses.append(status)
            rows[GOAL_RESULT].append([
                rid, str(row_no), g.get("stable_id") or "", source_kind(g.get("source_label")),
                g.get("template_code") or "", g.get("template_structure") or "", g.get("resolved_structure") or "",
                g.get("mapping_status") or "unresolved", g.get("objective") or "", g.get("comparator") or "",
                number(g.get("goal_value")), number(g.get("variation_value")), number(g.get("achieved_value")),
                g.get("unit") or "", status, "", g.get("source_label") or "",
            ])
        goal_sha = goal_table_sha256(content.goal_rows)
        for p in content.parameter_rows:
            value = p.get("value")
            available = number(value) != ""
            rows[PLAN_PARAMETER].append([
                rid, str(len(rows[PLAN_PARAMETER]) + 1), p["scope"], p.get("entity") or "", p["name"],
                number(value), p.get("unit") or "", p.get("method") or "",
                "available" if available else "unavailable", text(p.get("note")),
            ])
        for row_no, c in enumerate(content.check_rows, start=1):
            status = c.get("status") if c.get("status") in CHECK_STATUSES else "not_evaluated"
            checks += 1
            if status in OPEN_CHECK_STATUSES:
                open_findings += 1
            unit = c.get("unit") or ""
            numeric = unit not in ("text", "boolean")
            observed_raw, expected_raw = c.get("observed_value"), c.get("expected_value")
            rows[CHECK_FINDING].append([
                rid, str(row_no), c.get("check_code") or "", c.get("category") or "", c.get("severity") or "",
                status, "1", text(observed_raw), text(expected_raw),
                number(try_parse_invariant_number(observed_raw)) if numeric else "",
                number(try_parse_invariant_number(expected_raw)) if numeric else "",
                unit, text(c.get("message")),
            ])

    rs_sha = ruleset_sha256(context.configuration_inputs, goal_sha)

    def count(status):
        return "" if content is None else str(sum(1 for s in statuses if s == status))

    finished = context.finished_utc or datetime.now(timezone.utc)
    rows[REVIEW].append([
        str(SCHEMA_VERSION), rid, context.mode, utc(context.started_utc), utc(finished),
        context.clearplan_version or "", context.build_id or "", pseudonymizer.identifier_mode,
        pseudonymizer.user(context.user_id), pseudonymizer.patient(context.patient_id),
        pseudonymizer.course(context.patient_id, context.course_id),
        pseudonymizer.plan(context.plan_uid, context.patient_id, context.course_id, context.plan_id),
        (context.plan_id or "") if (options.log_plan_label or pseudonymizer.clear_identifiers) else "",
        context.planning_item_type or "", utc(context.plan_created_utc), context.plan_approval_status or "",
        utc(context.planning_approval_utc), utc(context.treatment_approval_utc),
        "" if content is None else _distinct(content.machine_ids),
        "" if content is None else _distinct(content.techniques),
        "" if content is None else _distinct(content.mlc_models),
        "" if content is None else number(content.total_dose_gy),
        "" if content is None else number(content.dose_per_fraction_gy),
        "" if content is None else integer(content.fractions),
        context.goal_table_id or "", goal_sha, rs_sha,
        "" if content is None else str(len(statuses)),
        count("goal"), count("variation"), count("not_met"), count("not_evaluated"),
        "" if content is None else str(checks),
        "" if content is None else integer(context.n_checks_excluded),
        "" if content is None else str(open_findings),
        flag(context.report_written), (context.report_sha256 or "").lower(), flag(context.report_filed),
        context.run_status or RUN_OK, scrubber.scrub(context.run_message),
    ])
    for table in WRITE_ORDER:
        for row in rows[table]:
            if len(row) != len(COLUMNS[table]):
                raise AssertionError("%s row has %d fields; expected %d" % (table, len(row), len(COLUMNS[table])))
    return rows


# --------------------------------------------------------------------------------------------
# Writer (ReviewLogWriter)
# --------------------------------------------------------------------------------------------

class ReviewLogSchemaError(IOError):
    pass


class _DirectoryLock(object):
    """Cross-process lock on ``.review-log.lock``.

    The C# writer opens the lock file with FileShare.None. On Windows an open Python handle therefore
    blocks the C# writer, and a C# holder makes Python's open fail (retried). Python-Python exclusion
    uses an additional byte-range lock (msvcrt) or flock (POSIX).
    """

    def __init__(self, directory, timeout_seconds):
        self._path = os.path.join(directory, LOCK_FILE_NAME)
        self._timeout = max(0.0, float(timeout_seconds))
        self._handle = None

    def __enter__(self):
        deadline = time.monotonic() + self._timeout
        while True:
            try:
                handle = open(self._path, "a+b")
            except (OSError, IOError):
                handle = None
            if handle is not None:
                if _try_lock(handle):
                    self._handle = handle
                    return self
                handle.close()
            if time.monotonic() >= deadline:
                raise TimeoutError("Review-log lock not acquired within %d ms." % int(self._timeout * 1000))
            time.sleep(0.02 + random.random() * 0.06)

    def __exit__(self, *exc):
        try:
            _unlock(self._handle)
        finally:
            self._handle.close()
        return False


def _try_lock(handle):
    try:
        if os.name == "nt":
            import msvcrt
            handle.seek(0)
            msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
        else:
            import fcntl
            fcntl.flock(handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        return True
    except (OSError, IOError):
        return False


def _unlock(handle):
    try:
        if os.name == "nt":
            import msvcrt
            handle.seek(0)
            msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)
        else:
            import fcntl
            fcntl.flock(handle.fileno(), fcntl.LOCK_UN)
    except (OSError, IOError):
        pass


def header_line(table):
    return format_line(COLUMNS[table])


class ReviewLogWriter(object):
    """Append-only writer for the five tables; refuses to append to a file with a different header."""

    def __init__(self, directory, lock_timeout_seconds=5.0):
        if not directory or str(directory).strip() == "":
            raise ValueError("A review-log directory is required.")
        self.directory = os.fspath(directory)
        self.lock_timeout_seconds = lock_timeout_seconds

    def path_for(self, table):
        return os.path.join(self.directory, table + ".csv")

    def write(self, rows):
        os.makedirs(self.directory, exist_ok=True)
        counts = {}
        with _DirectoryLock(self.directory, self.lock_timeout_seconds):
            for table in WRITE_ORDER:
                self._verify_header(table)
            for table in WRITE_ORDER:
                self._append(table, rows.get(table, []))
                counts[table] = len(rows.get(table, []))
        return counts

    def _verify_header(self, table):
        path = self.path_for(table)
        if not os.path.exists(path):
            return
        expected = header_line(table)
        with open(path, "rb") as handle:
            head = handle.read(len(expected.encode("utf-8")) + 3)
        if not head:
            return
        actual = head.decode("utf-8", errors="replace")
        if actual.startswith("\ufeff"):
            actual = actual[1:]
        if not actual.startswith(expected):
            raise ReviewLogSchemaError("Existing %s.csv has a different header; refusing to append schema version %d rows."
                                       % (table, SCHEMA_VERSION))

    def _append(self, table, rows):
        path = self.path_for(table)
        parts = []
        mode = "r+b" if os.path.exists(path) else "w+b"
        with open(path, mode) as handle:
            handle.seek(0, os.SEEK_END)
            size = handle.tell()
            if size == 0:
                parts.append(header_line(table))
            else:
                handle.seek(size - 1)
                if handle.read(1) != b"\n":
                    parts.append("\r\n")  # repair an interrupted last line
            for row in rows:
                if len(row) != len(COLUMNS[table]):
                    raise ValueError("%s row has %d fields; expected %d." % (table, len(row), len(COLUMNS[table])))
                parts.append(format_line(row))
            if not parts:
                return
            handle.seek(0, os.SEEK_END)
            handle.write("".join(parts).encode("utf-8"))
            handle.flush()
            os.fsync(handle.fileno())
