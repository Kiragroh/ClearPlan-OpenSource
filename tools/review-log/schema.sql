-- ClearPlan structured review log, schema version 1 (SQLite >= 3.25; window functions).
-- Source of truth for column names: ClearPlan.Core/ReviewLog/ReviewLogSchema.cs and docs/REVIEW_LOG.md.
-- Empty CSV fields are imported as NULL (= not available). NULL is never 0.
PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS review (
    schema_version          INTEGER NOT NULL,
    review_id               TEXT    NOT NULL PRIMARY KEY,
    mode                    TEXT    NOT NULL CHECK (mode IN ('live', 'headless', 'retrospective')),
    started_utc             TEXT,
    finished_utc            TEXT,
    clearplan_version       TEXT,
    build_id                TEXT,
    identifier_mode         TEXT    NOT NULL CHECK (identifier_mode IN ('hmac_sha256', 'clear')),
    user_key                TEXT,
    patient_key             TEXT,
    course_key              TEXT,
    plan_key                TEXT,
    plan_label              TEXT,
    planning_item_type      TEXT,
    plan_created_utc        TEXT,
    plan_approval_status    TEXT,
    planning_approval_utc   TEXT,
    treatment_approval_utc  TEXT,
    machine_id              TEXT,
    technique               TEXT,
    mlc_model               TEXT,
    total_dose_gy           REAL,
    dose_per_fraction_gy    REAL,
    fractions               INTEGER,
    goal_table_id           TEXT,
    goal_table_sha256       TEXT,
    ruleset_sha256          TEXT,
    n_goals                 INTEGER,
    n_goal                  INTEGER,
    n_variation             INTEGER,
    n_not_met               INTEGER,
    n_not_evaluated         INTEGER,
    n_checks                INTEGER,
    n_checks_excluded       INTEGER,
    n_findings_open         INTEGER,
    report_written          INTEGER CHECK (report_written IN (0, 1)),
    report_sha256           TEXT,
    report_filed            INTEGER CHECK (report_filed IN (0, 1)),
    run_status              TEXT    NOT NULL CHECK (run_status IN ('ok', 'partial', 'failed')),
    run_message             TEXT
);

CREATE TABLE IF NOT EXISTS goal_result (
    review_id           TEXT    NOT NULL REFERENCES review (review_id),
    row_no              INTEGER NOT NULL,
    stable_id           TEXT,
    source_kind         TEXT    CHECK (source_kind IN ('constraint_table', 'eclipse_clinical_goal', 'default_rule')),
    template_code       TEXT,
    template_structure  TEXT,
    resolved_structure  TEXT,
    mapping_status      TEXT    CHECK (mapping_status IN ('resolved', 'ambiguous', 'unresolved')),
    objective           TEXT,
    comparator          TEXT,
    goal_value          REAL,
    variation_value     REAL,
    achieved_value      REAL,
    unit                TEXT,
    status              TEXT    NOT NULL CHECK (status IN ('goal', 'variation', 'not_met', 'not_evaluated')),
    priority            TEXT,
    source_label        TEXT,
    PRIMARY KEY (review_id, row_no)
);

CREATE TABLE IF NOT EXISTS plan_parameter (
    review_id     TEXT    NOT NULL REFERENCES review (review_id),
    row_no        INTEGER NOT NULL,
    scope         TEXT    NOT NULL CHECK (scope IN ('plan', 'target', 'beam')),
    entity        TEXT,
    name          TEXT    NOT NULL,
    value         REAL,
    unit          TEXT,
    method        TEXT,
    availability  TEXT    NOT NULL CHECK (availability IN ('available', 'unavailable')),
    note          TEXT,
    PRIMARY KEY (review_id, row_no),
    CHECK ((availability = 'available') = (value IS NOT NULL))
);

CREATE TABLE IF NOT EXISTS check_finding (
    review_id         TEXT    NOT NULL REFERENCES review (review_id),
    row_no            INTEGER NOT NULL,
    check_code        TEXT,
    category          TEXT,
    severity          TEXT,
    status            TEXT    NOT NULL CHECK (status IN ('ok', 'variation', 'fail', 'info', 'not_evaluated')),
    included          INTEGER NOT NULL CHECK (included IN (0, 1)),
    observed_value    TEXT,
    expected_value    TEXT,
    observed_numeric  REAL,
    expected_numeric  REAL,
    unit              TEXT,
    message           TEXT,
    PRIMARY KEY (review_id, row_no)
);

CREATE TABLE IF NOT EXISTS ruleset_revision (
    review_id             TEXT    NOT NULL REFERENCES review (review_id),
    config_key            TEXT    NOT NULL,
    category              TEXT,
    revision_number       INTEGER,
    sha256                TEXT,
    revision_created_utc  TEXT,
    availability          TEXT    CHECK (availability IN ('available', 'unavailable', 'not_configured', 'not_hashed')),
    PRIMARY KEY (review_id, config_key)
);

CREATE INDEX IF NOT EXISTS ix_review_started   ON review (started_utc);
CREATE INDEX IF NOT EXISTS ix_review_plan      ON review (plan_key, started_utc);
CREATE INDEX IF NOT EXISTS ix_review_patient   ON review (patient_key);
CREATE INDEX IF NOT EXISTS ix_review_ruleset   ON review (ruleset_sha256);
CREATE INDEX IF NOT EXISTS ix_review_table     ON review (goal_table_id);
CREATE INDEX IF NOT EXISTS ix_goal_objective   ON goal_result (template_structure, objective, status);
CREATE INDEX IF NOT EXISTS ix_param_name       ON plan_parameter (name, scope);
CREATE INDEX IF NOT EXISTS ix_check_code       ON check_finding (check_code, status);
CREATE INDEX IF NOT EXISTS ix_ruleset_key      ON ruleset_revision (config_key, sha256);

-- ---------------------------------------------------------------------------------------------
-- Example views. All exclude reviews with run_status = 'failed' (no evaluated content).
-- Months are UTC calendar months (substr of ISO-8601 started_utc).
-- ---------------------------------------------------------------------------------------------

DROP VIEW IF EXISTS v_latest_review_per_plan;
CREATE VIEW v_latest_review_per_plan AS
SELECT * FROM (
    SELECT r.*, ROW_NUMBER() OVER (PARTITION BY r.plan_key, r.mode ORDER BY r.started_utc DESC, r.review_id) AS rn
    FROM review r WHERE r.run_status <> 'failed' AND r.plan_key IS NOT NULL
) WHERE rn = 1;

-- Check code x month: how often a check was open among the reviews in which it was evaluated (included = 1).
DROP VIEW IF EXISTS v_recurring_findings;
CREATE VIEW v_recurring_findings AS
SELECT c.check_code,
       substr(r.started_utc, 1, 7)                                                     AS month,
       COUNT(*)                                                                        AS n_evaluated,
       SUM(CASE WHEN c.status IN ('variation', 'fail', 'not_evaluated') THEN 1 ELSE 0 END) AS n_open,
       SUM(CASE WHEN c.status = 'fail' THEN 1 ELSE 0 END)                              AS n_fail,
       ROUND(1.0 * SUM(CASE WHEN c.status IN ('variation', 'fail', 'not_evaluated') THEN 1 ELSE 0 END) / COUNT(*), 4) AS open_share
FROM check_finding c JOIN review r ON r.review_id = c.review_id
WHERE c.included = 1 AND r.run_status <> 'failed'
GROUP BY c.check_code, substr(r.started_utc, 1, 7);

-- Goal table x structure x objective: status counts; shares use evaluated rows (not_evaluated excluded) as denominator.
DROP VIEW IF EXISTS v_goal_status_by_table;
CREATE VIEW v_goal_status_by_table AS
SELECT r.goal_table_id,
       g.source_kind,
       g.template_structure,
       g.objective,
       COUNT(*)                                                   AS n_rows,
       SUM(CASE WHEN g.status = 'goal' THEN 1 ELSE 0 END)          AS n_goal,
       SUM(CASE WHEN g.status = 'variation' THEN 1 ELSE 0 END)     AS n_variation,
       SUM(CASE WHEN g.status = 'not_met' THEN 1 ELSE 0 END)       AS n_not_met,
       SUM(CASE WHEN g.status = 'not_evaluated' THEN 1 ELSE 0 END) AS n_not_evaluated,
       ROUND(1.0 * SUM(CASE WHEN g.status = 'variation' THEN 1 ELSE 0 END) /
             NULLIF(SUM(CASE WHEN g.status <> 'not_evaluated' THEN 1 ELSE 0 END), 0), 4) AS variation_share,
       ROUND(1.0 * SUM(CASE WHEN g.status = 'not_met' THEN 1 ELSE 0 END) /
             NULLIF(SUM(CASE WHEN g.status <> 'not_evaluated' THEN 1 ELSE 0 END), 0), 4) AS not_met_share
FROM goal_result g JOIN review r ON r.review_id = g.review_id
WHERE r.run_status <> 'failed'
GROUP BY r.goal_table_id, g.source_kind, g.template_structure, g.objective;

-- Parameter x scope x month: n, median, nearest-rank quartiles (q1 = rank ceil(n/4), q3 = rank ceil(3n/4)).
DROP VIEW IF EXISTS v_parameter_trend;
CREATE VIEW v_parameter_trend AS
WITH v AS (
    SELECT p.name, p.scope, p.unit, substr(r.started_utc, 1, 7) AS month, p.value
    FROM plan_parameter p JOIN review r ON r.review_id = p.review_id
    WHERE p.availability = 'available' AND p.value IS NOT NULL AND r.run_status <> 'failed'
), ranked AS (
    SELECT v.*,
           ROW_NUMBER() OVER (PARTITION BY name, scope, month ORDER BY value) AS rk,
           COUNT(*)     OVER (PARTITION BY name, scope, month)                 AS n
    FROM v
)
SELECT name, scope, unit, month, n,
       AVG(CASE WHEN rk IN ((n + 1) / 2, (n + 2) / 2) THEN value END) AS median,
       MAX(CASE WHEN rk = (n + 3) / 4 THEN value END)                 AS q1,
       MAX(CASE WHEN rk = (3 * n + 3) / 4 THEN value END)             AS q3,
       MIN(value) AS min_value,
       MAX(value) AS max_value
FROM ranked
GROUP BY name, scope, unit, month, n;

-- Rule-set epochs: one row per combined rule-set fingerprint.
DROP VIEW IF EXISTS v_ruleset_epoch;
CREATE VIEW v_ruleset_epoch AS
SELECT ruleset_sha256, MIN(started_utc) AS first_seen_utc, MAX(started_utc) AS last_seen_utc, COUNT(*) AS n_reviews
FROM review WHERE run_status <> 'failed' AND ruleset_sha256 IS NOT NULL
GROUP BY ruleset_sha256;

-- Configuration documents that differ between consecutive rule-set epochs.
DROP VIEW IF EXISTS v_ruleset_config;
CREATE VIEW v_ruleset_config AS
SELECT DISTINCT r.ruleset_sha256, rr.config_key, rr.category, rr.sha256, rr.revision_number, rr.availability
FROM ruleset_revision rr JOIN review r ON r.review_id = rr.review_id;

-- Check code x rule-set epoch, with the previous epoch's open share for before/after comparison.
DROP VIEW IF EXISTS v_rule_change;
CREATE VIEW v_rule_change AS
WITH per AS (
    SELECT c.check_code, r.ruleset_sha256,
           COUNT(*) AS n_evaluated,
           SUM(CASE WHEN c.status IN ('variation', 'fail', 'not_evaluated') THEN 1 ELSE 0 END) AS n_open
    FROM check_finding c JOIN review r ON r.review_id = c.review_id
    WHERE c.included = 1 AND r.run_status <> 'failed' AND r.ruleset_sha256 IS NOT NULL
    GROUP BY c.check_code, r.ruleset_sha256
), e AS (
    SELECT per.*, ep.first_seen_utc FROM per JOIN v_ruleset_epoch ep ON ep.ruleset_sha256 = per.ruleset_sha256
)
SELECT check_code, ruleset_sha256, first_seen_utc, n_evaluated, n_open,
       ROUND(1.0 * n_open / n_evaluated, 4) AS open_share,
       LAG(ruleset_sha256) OVER (PARTITION BY check_code ORDER BY first_seen_utc)                       AS previous_ruleset_sha256,
       LAG(ROUND(1.0 * n_open / n_evaluated, 4)) OVER (PARTITION BY check_code ORDER BY first_seen_utc) AS previous_open_share
FROM e;
