#!/usr/bin/env python3
"""Import ClearPlan structured review-log CSV files into SQLite.

Standard library only (csv, sqlite3). Usage:

    python import_review_log.py --log-dir <dir> [--log-dir <dir2> ...] --db review_log.sqlite
                                [--schema schema.sql] [--export-views <dir>]

* Creates the schema (tables, indices, example views) from schema.sql if needed.
* Verifies every CSV header against the SQLite table definition (exact names and order).
* Empty fields become NULL (= not available); numeric columns are parsed strictly (no locale).
* Idempotent: rows are keyed (review_id + row_no / config_key); re-importing the same files adds nothing.
* Child rows whose review row is absent (interrupted write) are skipped and counted as orphans.
Prints a JSON summary to stdout. Exit code 0 on success, 1 on validation errors.
"""
import argparse
import csv
import json
import os
import sqlite3
import sys

SCHEMA_VERSION = 1
TABLES = ["review", "goal_result", "plan_parameter", "check_finding", "ruleset_revision"]
EXAMPLE_VIEWS = ["v_recurring_findings", "v_goal_status_by_table", "v_parameter_trend", "v_rule_change",
                 "v_ruleset_epoch", "v_ruleset_config", "v_latest_review_per_plan"]


class ImportError_(Exception):
    pass


def table_columns(conn, table):
    return [(row[1], (row[2] or "").upper()) for row in conn.execute("PRAGMA table_info(%s)" % table)]


def convert(value, sql_type, table, column, line):
    if value == "":
        return None
    if sql_type == "INTEGER":
        try:
            return int(value)
        except ValueError:
            raise ImportError_("%s.csv line %d: column %s expects an integer, found %r" % (table, line, column, value))
    if sql_type == "REAL":
        try:
            number = float(value)
        except ValueError:
            raise ImportError_("%s.csv line %d: column %s expects a number, found %r" % (table, line, column, value))
        if number != number or number in (float("inf"), float("-inf")):
            raise ImportError_("%s.csv line %d: column %s is not finite" % (table, line, column))
        return number
    return value


def read_table(path, table, columns):
    with open(path, "r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.reader(handle)
        try:
            header = next(reader)
        except StopIteration:
            return []
        expected = [name for name, _ in columns]
        if header != expected:
            raise ImportError_("%s: header mismatch.\n  expected: %s\n  found:    %s" % (path, ",".join(expected), ",".join(header)))
        rows = []
        for line_no, row in enumerate(reader, start=2):
            if not row:
                continue
            if len(row) != len(columns):
                raise ImportError_("%s line %d: %d fields, expected %d" % (path, line_no, len(row), len(columns)))
            rows.append([convert(v, t, table, n, line_no) for v, (n, t) in zip(row, columns)])
        return rows


def import_dirs(conn, log_dirs):
    summary = {"schema_version": SCHEMA_VERSION, "inserted": {}, "skipped_duplicates": {}, "orphans": {}, "files": []}
    for table in TABLES:
        summary["inserted"][table] = 0
        summary["skipped_duplicates"][table] = 0
        summary["orphans"][table] = 0
    for log_dir in log_dirs:
        data = {}
        for table in TABLES:
            path = os.path.join(log_dir, table + ".csv")
            if not os.path.exists(path):
                data[table] = []
                continue
            data[table] = read_table(path, table, table_columns(conn, table))
            summary["files"].append(path)
        for row in data["review"]:
            if row[0] != SCHEMA_VERSION:
                raise ImportError_("review %s has schema_version %r; this importer handles %d" % (row[1], row[0], SCHEMA_VERSION))
        with conn:
            for table in ["review"] + [t for t in TABLES if t != "review"]:
                columns = table_columns(conn, table)
                placeholders = ",".join("?" for _ in columns)
                sql = "INSERT OR IGNORE INTO %s (%s) VALUES (%s)" % (table, ",".join(n for n, _ in columns), placeholders)
                review_index = [n for n, _ in columns].index("review_id")
                for row in data[table]:
                    if table != "review":
                        known = conn.execute("SELECT 1 FROM review WHERE review_id = ?", (row[review_index],)).fetchone()
                        if known is None:
                            summary["orphans"][table] += 1
                            continue
                    cursor = conn.execute(sql, row)
                    if cursor.rowcount == 1:
                        summary["inserted"][table] += 1
                    else:
                        summary["skipped_duplicates"][table] += 1
    summary["totals"] = {t: conn.execute("SELECT COUNT(*) FROM %s" % t).fetchone()[0] for t in TABLES}
    return summary


def export_views(conn, out_dir):
    os.makedirs(out_dir, exist_ok=True)
    written = {}
    for view in EXAMPLE_VIEWS:
        cursor = conn.execute("SELECT * FROM %s" % view)
        names = [d[0] for d in cursor.description]
        rows = cursor.fetchall()
        with open(os.path.join(out_dir, view + ".csv"), "w", encoding="utf-8", newline="") as handle:
            writer = csv.writer(handle, lineterminator="\r\n")
            writer.writerow(names)
            writer.writerows(rows)
        written[view] = len(rows)
    return written


def main(argv=None):
    here = os.path.dirname(os.path.abspath(__file__))
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--log-dir", action="append", required=True, help="review-log directory (repeatable)")
    parser.add_argument("--db", required=True, help="SQLite database file (created if absent)")
    parser.add_argument("--schema", default=os.path.join(here, "schema.sql"))
    parser.add_argument("--export-views", help="write each example view as CSV into this directory")
    args = parser.parse_args(argv)
    conn = sqlite3.connect(args.db)
    try:
        with open(args.schema, "r", encoding="utf-8") as handle:
            conn.executescript(handle.read())
        conn.execute("PRAGMA foreign_keys = ON")
        summary = import_dirs(conn, args.log_dir)
        if args.export_views:
            summary["views"] = export_views(conn, args.export_views)
        summary["sqlite_version"] = sqlite3.sqlite_version
        print(json.dumps(summary, indent=2, sort_keys=True))
        return 0
    except ImportError_ as error:
        print("import failed: %s" % error, file=sys.stderr)
        return 1
    finally:
        conn.close()


if __name__ == "__main__":
    sys.exit(main())
