"""TPS-independent core of the ClearPlan RayStation example.

UNVALIDATED EXAMPLE - NOT FOR CLINICAL USE.

Nothing in this package imports or calls a treatment-planning-system API. The
RayStation adapter (``raystation_adapter.py``) converts RayStation objects into
the neutral :class:`clearplan_review_core.model.PlanData`; everything else
(goal tables, structure mapping, evaluation, checks, review-log rows) is plain
Python standard library code and is shared with the tests.
"""

__all__ = [
    "evaluator",
    "objectives",
    "goal_table",
    "mapping",
    "model",
    "plan_metrics",
    "checks",
    "review_log",
    "engine",
    "html_summary",
]

VERSION = "0.1.0"
IMPLEMENTATION = "clearplan-raystation-example"
