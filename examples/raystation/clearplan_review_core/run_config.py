"""Run configuration shared by the live and retrospective entry scripts (JSON, ClearPlan.raystation_run.v1).

{
  "schema": "ClearPlan.raystation_run.v1",
  "goal_table": "rules/prostate_5x5.json",          required (CSV or JSON)
  "goal_table_id": null,                            optional override of the table name
  "structure_aliases": "rules/aliases.json",        optional (ClearPlan.structure_aliases.v1)
  "review_checks": "rules/review_checks.json",      optional (ClearPlan.review_checks.v1)
  "log_dir": "D:/ClearPlanLog/review_log",          required
  "key_file": "D:/ClearPlanState/review-log.key",   required unless log_clear_identifiers (same file as Eclipse)
  "log_clear_identifiers": false, "log_plan_label": false, "log_messages": true,
  "html_summary_dir": null,                          optional, live mode only
  "mining_list": "lists/mining.csv"                  optional, retrospective mode
}
Relative paths resolve against the directory of the run-configuration file.
"""

import json
import os

from .review_log import ReviewLogOptions

SCHEMA = "ClearPlan.raystation_run.v1"
ENVIRONMENT_VARIABLE = "CLEARPLAN_RS_RUN_CONFIG"


class RunConfigError(ValueError):
    pass


def _path(base, value):
    if value is None or str(value).strip() == "":
        return None
    value = os.path.expandvars(str(value))
    return os.path.normpath(value if os.path.isabs(value) else os.path.join(base, value))


def load(path=None, default_dir=None):
    if path is None:
        path = os.environ.get(ENVIRONMENT_VARIABLE)
    if path is None and default_dir is not None:
        candidate = os.path.join(default_dir, "clearplan_rs_run.json")
        path = candidate if os.path.exists(candidate) else None
    if path is None:
        raise RunConfigError("No run configuration: pass a path, set %s or place clearplan_rs_run.json "
                             "next to the script." % ENVIRONMENT_VARIABLE)
    with open(path, "rb") as handle:
        try:
            document = json.loads(handle.read().decode("utf-8-sig"))
        except ValueError as error:
            raise RunConfigError("Run configuration is malformed: %s" % error)
    if not isinstance(document, dict) or document.get("schema") != SCHEMA:
        raise RunConfigError("Run configuration must declare schema %r." % SCHEMA)
    base = os.path.dirname(os.path.abspath(path))
    config = {
        "path": os.path.abspath(path),
        "goal_table": _path(base, document.get("goal_table")),
        "goal_table_id": document.get("goal_table_id"),
        "structure_aliases": _path(base, document.get("structure_aliases")),
        "review_checks": _path(base, document.get("review_checks")),
        "log_dir": _path(base, document.get("log_dir")),
        "key_file": _path(base, document.get("key_file")),
        "html_summary_dir": _path(base, document.get("html_summary_dir")),
        "mining_list": _path(base, document.get("mining_list")),
    }
    for flag_name, default in (("log_clear_identifiers", False), ("log_plan_label", False), ("log_messages", True)):
        value = document.get(flag_name, default)
        if not isinstance(value, bool):
            raise RunConfigError("%s must be true or false." % flag_name)
        config[flag_name] = value
    if not config["goal_table"]:
        raise RunConfigError("goal_table is required.")
    if not config["log_dir"]:
        raise RunConfigError("log_dir is required.")
    if not config["log_clear_identifiers"]:
        if not config["key_file"]:
            raise RunConfigError("key_file is required for pseudonymised logging.")
        log_dir = os.path.normcase(os.path.abspath(config["log_dir"])) + os.sep
        if os.path.normcase(os.path.abspath(config["key_file"])).startswith(log_dir):
            raise RunConfigError("key_file must not be stored inside log_dir.")
    return config


def options(config):
    return ReviewLogOptions(log_clear_identifiers=config["log_clear_identifiers"],
                            log_plan_label=config["log_plan_label"], log_messages=config["log_messages"])
