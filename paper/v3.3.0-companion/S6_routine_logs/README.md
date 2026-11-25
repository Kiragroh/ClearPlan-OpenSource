# Supplement S6: Routine-use log analysis of the technical-check component (PlanCheck)

## Source and scope

Before it was integrated into ClearPlan, the technical-check component PlanCheck ran as a standalone script for all plans in our department. Every run logged its findings. A local viewer combined these logs into three tables: plans, findings that were absent at a later run, and findings still present at the last run. This supplement aggregates those tables for plans first checked between 3 July 2023 and 28 December 2024.

Only aggregated counts are given here. Patient, course and plan identifiers and free-text messages are not distributed. Messages appear only as templates, with numbers and quoted names removed, so that each error code has a label.

## Definitions (taken from the viewer code)

| Term | Definition |
| --- | --- |
| Included plan | Patient/course/plan with at least two PlanCheck runs by at least two different users. Plans checked only once are not in the export, so the total number of checked plans is not known. |
| Major / minor | PlanCheck classes "1 – Deviation" and "2 – Variation". These are not the ClearPlan goal states Goal/Variation. |
| Absent at a later run | A finding present in one run and missing from the next run of the same plan. This does not establish that an error was corrected: the plan, the check or its configuration may have changed. |
| Present at the last run | A finding still present in the last logged run. It may have been accepted deliberately; the log records no reason. |
| Excluded codes | 27 codes were switched off in the viewer (`settings.json`) and are not counted. |
| Category | Error codes grouped by hand from their message templates (`CATEGORY` in the script). |

**Correction relative to the viewer output:** the viewer writes the "absent at a later run" table before it removes plans with a single user. Of its 437 rows, 57 belong to single-user plans. They are excluded here, which leaves 380.

## Files

- `summary.json`: period, plans, patients, runs, users per plan, and totals by class.
- `by_category.csv`: per category:
  - plans with a finding in any run, all and major;
  - findings absent at a later run, major and minor;
  - findings present at the last run, major and minor.
- `by_code.csv`: per error code: category, class, counts, the first month with a finding absent at a later run, and the message template.
- `by_month.csv`: findings absent at a later run, by month, with the number of plans first checked in that month.
- `aggregate_plancheck.py`: the script that produced these files from the viewer export.

## Caveats

- **Checks changed over time.** They were added and modified during the period, and activation dates were not logged. Frequencies therefore cannot be compared over time or between codes.
- **Unexplained peak.** In April 2024, 95 findings were absent at a later run; other months had 8–31. The cause is not known and is not interpreted.
- **No comparison cohort.** PlanCheck was used for all plans, so there are no plans reviewed without it. The counts describe use, not detection rates or prevented errors.
- **One centre, one TPS.**

## Reproduction

```
python -I aggregate_plancheck.py <folder with PatientData_PlanCheck.csv, FoundFixes.csv, finalerrors.csv, settings.json> <output folder>
```

Requires Python 3.13 and pandas. The input files contain patient identifiers and remain in the department.
