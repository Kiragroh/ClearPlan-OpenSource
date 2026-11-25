using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using ClearPlan.Core.PlanAnalysis;
using ClearPlan.Core.Review;
using ClearPlan.Core.ReviewLog;
using ClearPlan.Core.Settings;
using ClearPlan.Core.Simulation;

namespace ClearPlan.Core.Tests
{
    internal static class ReviewLogTests
    {
        // Test-only key (bytes 0..31). Never a production secret.
        private static readonly byte[] SyntheticKey = Enumerable.Range(0, 32).Select(i => (byte)i).ToArray();

        public static void SchemaHeadersAreStable()
        {
            TestAssert.Equal(1, ReviewLogSchema.SchemaVersion);
            TestAssert.Equal("schema_version,review_id,mode,started_utc,finished_utc,clearplan_version,build_id,identifier_mode,user_key,patient_key,course_key,plan_key,plan_label,planning_item_type,plan_created_utc,plan_approval_status,planning_approval_utc,treatment_approval_utc,machine_id,technique,mlc_model,total_dose_gy,dose_per_fraction_gy,fractions,goal_table_id,goal_table_sha256,ruleset_sha256,n_goals,n_goal,n_variation,n_not_met,n_not_evaluated,n_checks,n_checks_excluded,n_findings_open,report_written,report_sha256,report_filed,run_status,run_message",
                string.Join(",", ReviewLogSchema.ReviewColumns));
            TestAssert.Equal("review_id,row_no,stable_id,source_kind,template_code,template_structure,resolved_structure,mapping_status,objective,comparator,goal_value,variation_value,achieved_value,unit,status,priority,source_label",
                string.Join(",", ReviewLogSchema.GoalResultColumns));
            TestAssert.Equal("review_id,row_no,scope,entity,name,value,unit,method,availability,note", string.Join(",", ReviewLogSchema.PlanParameterColumns));
            TestAssert.Equal("review_id,row_no,check_code,category,severity,status,included,observed_value,expected_value,observed_numeric,expected_numeric,unit,message",
                string.Join(",", ReviewLogSchema.CheckFindingColumns));
            TestAssert.Equal("review_id,config_key,category,revision_number,sha256,revision_created_utc,availability", string.Join(",", ReviewLogSchema.RulesetRevisionColumns));
            TestAssert.Equal(ReviewLogSchema.ReviewTable, ReviewLogSchema.WriteOrder.Last(), "The review row must be written last.");
            foreach (string table in ReviewLogSchema.WriteOrder)
                foreach (string column in ReviewLogSchema.Columns(table))
                    TestAssert.True(System.Text.RegularExpressions.Regex.IsMatch(column, "^[a-z][a-z0-9_]*$"), "Column is not snake_case: " + column);
            // schema.sql must declare exactly these columns, in this order.
            string sql = File.ReadAllText(Path.Combine("tools", "review-log", "schema.sql"));
            foreach (string table in ReviewLogSchema.WriteOrder)
            {
                int start = sql.IndexOf("CREATE TABLE IF NOT EXISTS " + table + " (", StringComparison.Ordinal);
                TestAssert.True(start >= 0, "schema.sql lacks table " + table);
                int end = sql.IndexOf("\n);", start, StringComparison.Ordinal);
                var declared = sql.Substring(start, end - start).Split('\n').Skip(1).Select(l => l.Trim())
                    .Where(l => l.Length > 0 && !l.StartsWith("PRIMARY KEY", StringComparison.Ordinal) && !l.StartsWith("CHECK", StringComparison.Ordinal))
                    .Select(l => l.Split(' ')[0]).ToList();
                TestAssert.Equal(string.Join(",", ReviewLogSchema.Columns(table)), string.Join(",", declared), "schema.sql column drift in " + table);
            }
        }

        public static void CsvEscapingRoundTrips()
        {
            var values = new[] { "plain", "comma,inside", "quote\"inside", "line\nbreak", "crlf\r\nbreak", "semi;colon", " leading", "trailing ",
                "Unicode: ä ö ü ß µ ° ± ≤ ≥ 中文 🙂", "", "D0.035cc[Gy]", "\"fully quoted\"" };
            string line = ReviewLogCsv.FormatLine(values);
            TestAssert.True(line.EndsWith("\r\n", StringComparison.Ordinal));
            TestAssert.Equal("semi;colon", ReviewLogCsv.Escape("semi;colon"), "Semicolons are data, not delimiters.");
            TestAssert.Equal("\"comma,inside\"", ReviewLogCsv.Escape("comma,inside"));
            TestAssert.Equal("\"quote\"\"inside\"", ReviewLogCsv.Escape("quote\"inside"));
            List<List<string>> parsed = ReviewLogCsv.Parse(line + line);
            TestAssert.Equal(2, parsed.Count);
            TestAssert.Equal(string.Join("|", values), string.Join("|", parsed[0]));
            TestAssert.Equal("1.5", ReviewLogCsv.Number(1.5));
            TestAssert.Equal(string.Empty, ReviewLogCsv.Number(double.NaN));
            TestAssert.Equal(string.Empty, ReviewLogCsv.Number(double.PositiveInfinity));
            TestAssert.Equal(string.Empty, ReviewLogCsv.Number(null));
            TestAssert.Equal("2026-01-02T03:04:05.006Z", ReviewLogCsv.Utc(new DateTimeOffset(2026, 1, 2, 4, 4, 5, 6, TimeSpan.FromHours(1))));
            var previous = CultureInfo.CurrentCulture;
            try
            {
                CultureInfo.CurrentCulture = new CultureInfo("de-DE");
                TestAssert.Equal("24.2", ReviewLogCsv.Number(24.2), "Numbers must be invariant culture.");
            }
            finally { CultureInfo.CurrentCulture = previous; }
        }

        public static void NotEvaluatedIsNeverWrittenAsPassed()
        {
            foreach (string raw in new[] { null, "", "not-evaluated", "info", "unknown", "available", "fallback", "Goal", "met" })
                TestAssert.Equal(ReviewLogGoalStatus.NotEvaluated, ReviewLogGoalStatus.FromSnapshot(raw), "Status '" + raw + "' must not pass.");
            TestAssert.Equal(ReviewLogGoalStatus.Goal, ReviewLogGoalStatus.FromSnapshot("pass"));
            TestAssert.Equal(ReviewLogGoalStatus.Variation, ReviewLogGoalStatus.FromSnapshot("variation"));
            TestAssert.Equal(ReviewLogGoalStatus.NotMet, ReviewLogGoalStatus.FromSnapshot("fail"));
            TestAssert.Equal(ReviewLogCheckStatus.NotEvaluated, ReviewLogCheckStatus.FromSnapshot("weird"));
            var snapshot = MinimalSnapshot();
            snapshot.PqmRows.Add(new ReviewPqmRow { StableId = "a", Objective = "D95%[Gy]", Status = "not-evaluated", AchievedValue = 59.1, Unit = "Gy" });
            snapshot.PqmRows.Add(new ReviewPqmRow { StableId = "b", Objective = "Dmean[Gy]", Status = null, Unit = "Gy" });
            snapshot.PqmRows.Add(new ReviewPqmRow { StableId = "c", Objective = "D0.035cc[Gy]", Status = "pass", AchievedValue = 44.0, Goal = 45, Unit = "Gy" });
            var record = Build(snapshot, Context(ReviewLogModes.Live));
            var goals = record.Rows[ReviewLogSchema.GoalResultTable];
            int status = Index(ReviewLogSchema.GoalResultColumns, "status");
            TestAssert.Equal("not_evaluated,not_evaluated,goal", string.Join(",", goals.Select(g => g[status])));
            string[] review = record.Rows[ReviewLogSchema.ReviewTable].Single();
            TestAssert.Equal("1", review[Index(ReviewLogSchema.ReviewColumns, "n_goal")]);
            TestAssert.Equal("2", review[Index(ReviewLogSchema.ReviewColumns, "n_not_evaluated")]);
        }

        public static void UnavailableValuesStayEmpty()
        {
            var snapshot = MinimalSnapshot();
            snapshot.PqmRows.Add(new ReviewPqmRow { StableId = "u", Objective = "V20Gy[%]", Status = "not-evaluated", Goal = double.NaN, Variation = null, AchievedValue = null, Unit = "%" });
            snapshot.PlanAnalysis = new ReviewPlanAnalysis { TotalMetersetMu = null, MuPerGy = double.PositiveInfinity, Pam = null, PamReason = "No target." };
            snapshot.PlanAnalysis.TargetQuality.Add(new ReviewTargetQuality { StructureId = "PTV", PaddickCi = null, GradientIndex = 3.2 });
            snapshot.Plans[0].TotalDoseGy = null;
            var record = Build(snapshot, Context(ReviewLogModes.Headless));
            string[] goal = record.Rows[ReviewLogSchema.GoalResultTable].Single();
            foreach (string column in new[] { "goal_value", "variation_value", "achieved_value" })
                TestAssert.Equal(string.Empty, goal[Index(ReviewLogSchema.GoalResultColumns, column)], column + " must be empty, not 0.");
            var parameters = record.Rows[ReviewLogSchema.PlanParameterTable];
            int name = Index(ReviewLogSchema.PlanParameterColumns, "name"), value = Index(ReviewLogSchema.PlanParameterColumns, "value"),
                availability = Index(ReviewLogSchema.PlanParameterColumns, "availability"), scope = Index(ReviewLogSchema.PlanParameterColumns, "scope");
            foreach (string[] row in parameters)
            {
                TestAssert.Equal(row[value].Length > 0 ? "available" : "unavailable", row[availability]);
                TestAssert.True(row[value] != "0" || row[availability] == "available");
            }
            TestAssert.Equal("unavailable", parameters.Single(r => r[scope] == "plan" && r[name] == "pam")[availability]);
            TestAssert.Equal("unavailable", parameters.Single(r => r[scope] == "plan" && r[name] == "mu_per_gy")[availability]);
            TestAssert.Equal("3.2", parameters.Single(r => r[name] == "gradient_index")[value]);
            TestAssert.Equal(string.Empty, record.Rows[ReviewLogSchema.ReviewTable].Single()[Index(ReviewLogSchema.ReviewColumns, "total_dose_gy")]);
        }

        public static void PseudonymisationIsKeyedAndDeterministic()
        {
            var a = new ReviewLogPseudonymizer(SyntheticKey, false);
            var b = new ReviewLogPseudonymizer(SyntheticKey, false);
            var other = new ReviewLogPseudonymizer(Enumerable.Range(1, 32).Select(i => (byte)i).ToArray(), false);
            string key = a.Patient("SYNTH-PAT-001");
            TestAssert.True(System.Text.RegularExpressions.Regex.IsMatch(key, "^pat_[0-9a-f]{32}$"), key);
            TestAssert.Equal(key, b.Patient("SYNTH-PAT-001"), "Same key must give the same pseudonym.");
            TestAssert.Equal(key, a.Patient(" synth-pat-001 "), "Patient IDs are normalised (trim, case).");
            TestAssert.False(key == other.Patient("SYNTH-PAT-001"), "A different key must give a different pseudonym.");
            TestAssert.False(key.Contains("SYNTH") || key.Contains("001"));
            // Keyed HMAC, not a plain (dictionary-reversible) hash of the identifier.
            string plain;
            using (var sha = SHA256.Create()) plain = string.Concat(sha.ComputeHash(Encoding.UTF8.GetBytes("SYNTH-PAT-001")).Select(x => x.ToString("x2")));
            TestAssert.False(plain.StartsWith(key.Substring(4), StringComparison.Ordinal));
            string expected;
            using (var hmac = new HMACSHA256(SyntheticKey))
                expected = "pat_" + string.Concat(hmac.ComputeHash(Encoding.UTF8.GetBytes("pat\u001eSYNTH-PAT-001")).Take(16).Select(x => x.ToString("x2")));
            TestAssert.Equal(expected, key);
            TestAssert.False(a.Course("P1", "C1") == a.Course("P2", "C1"), "Course keys are patient-scoped.");
            TestAssert.Equal(a.Plan("1.2.3", "P1", "C1", "X"), a.Plan("1.2.3", "P9", "C9", "Y"), "Plan keys derive from the UID.");
            TestAssert.Equal(a.User("DOMAIN\\synthuser"), a.User("SYNTHUSER"));
            TestAssert.Equal(string.Empty, a.Patient(null));
            var clear = new ReviewLogPseudonymizer(null, true);
            TestAssert.Equal("SYNTH-PAT-001", clear.Patient("SYNTH-PAT-001"));
            TestAssert.Equal("clear", clear.IdentifierMode);
            TestAssert.Throws<ArgumentException>(() => new ReviewLogPseudonymizer(new byte[8], false));
            string directory = TempDirectory();
            try
            {
                string keyFile = Path.Combine(directory, "state", "review-log.key");
                byte[] first = ReviewLogPseudonymizer.LoadOrCreateKey(keyFile);
                byte[] second = ReviewLogPseudonymizer.LoadOrCreateKey(keyFile);
                TestAssert.Equal(32, first.Length);
                TestAssert.True(first.SequenceEqual(second), "The key file is created once and reused.");
                var parallel = Enumerable.Range(0, 8).AsParallel().Select(_ => Convert.ToBase64String(ReviewLogPseudonymizer.LoadOrCreateKey(keyFile))).Distinct().ToList();
                TestAssert.Equal(1, parallel.Count);
            }
            finally { Directory.Delete(directory, true); }
        }

        public static void OkChecksAreLoggedForDenominators()
        {
            var snapshot = MinimalSnapshot();
            snapshot.PlanCheckRows.Add(new ReviewCheckRow { CheckCode = "ct-age-days", Category = "Default", Status = "pass", Severity = "none", Unit = "count", ObservedValue = "12", ExpectedValue = "60", Message = "OK" });
            snapshot.PlanCheckRows.Add(new ReviewCheckRow { CheckCode = "hu-table", Category = "Default", Status = "fail", Severity = "error", Unit = "text", Message = "Mismatch" });
            snapshot.PlanCheckRows.Add(new ReviewCheckRow { CheckCode = "grid", Category = "Default", Status = "not-evaluated", Severity = "warning", Unit = "text" });
            snapshot.DisabledCheckCount = 1;
            var context = Context(ReviewLogModes.Live);
            context.ExcludedCheckRows.Add(new ReviewCheckRow { CheckCode = "PC-legacy", Category = "PlanCheck", Status = "fail", Unit = "text" });
            var record = Build(snapshot, context);
            var checks = record.Rows[ReviewLogSchema.CheckFindingTable];
            int status = Index(ReviewLogSchema.CheckFindingColumns, "status"), included = Index(ReviewLogSchema.CheckFindingColumns, "included");
            TestAssert.Equal("ok:1,fail:1,not_evaluated:1,fail:0", string.Join(",", checks.Select(c => c[status] + ":" + c[included])));
            TestAssert.Equal("12", checks[0][Index(ReviewLogSchema.CheckFindingColumns, "observed_numeric")]);
            TestAssert.Equal(string.Empty, checks[1][Index(ReviewLogSchema.CheckFindingColumns, "observed_numeric")]);
            string[] review = record.Rows[ReviewLogSchema.ReviewTable].Single();
            TestAssert.Equal("3", review[Index(ReviewLogSchema.ReviewColumns, "n_checks")]);
            TestAssert.Equal("1", review[Index(ReviewLogSchema.ReviewColumns, "n_checks_excluded")]);
            TestAssert.Equal("2", review[Index(ReviewLogSchema.ReviewColumns, "n_findings_open")]);
        }

        public static void FreeTextIsScrubbedAndLabelsOmitted()
        {
            var snapshot = MinimalSnapshot();
            snapshot.PlanCheckRows.Add(new ReviewCheckRow { CheckCode = "approval", Status = "variation", Unit = "text", Message = "Approved by Synthuser Example for MUSTERMANN, Erika (SYNTH-PAT-001)" });
            var context = Context(ReviewLogModes.Live);
            context.SensitiveTerms.AddRange(new[] { "SYNTH-PAT-001", "Mustermann", "Erika", "Synthuser Example" });
            context.RunMessage = "note for mustermann";
            var record = Build(snapshot, context);
            string message = record.Rows[ReviewLogSchema.CheckFindingTable].Single()[Index(ReviewLogSchema.CheckFindingColumns, "message")];
            TestAssert.Equal("Approved by [redacted] for [redacted], [redacted] ([redacted])", message);
            string[] review = record.Rows[ReviewLogSchema.ReviewTable].Single();
            TestAssert.Equal("note for [redacted]", review[Index(ReviewLogSchema.ReviewColumns, "run_message")]);
            TestAssert.Equal("[redacted] [redacted]", new ReviewLogTextScrubber(new[] { "Red", "dact" }).Scrub("red dact"), "Replacement text is never re-scanned.");
            TestAssert.Equal(string.Empty, review[Index(ReviewLogSchema.ReviewColumns, "plan_label")], "Plan IDs are not logged by default.");
            var labelled = ReviewLogRecordBuilder.Build(snapshot, context, new ReviewLogPseudonymizer(SyntheticKey, false), new ReviewLogOptions { LogPlanLabel = true, LogMessages = false });
            TestAssert.Equal("SYNTH_PLAN", labelled.Rows[ReviewLogSchema.ReviewTable].Single()[Index(ReviewLogSchema.ReviewColumns, "plan_label")]);
            TestAssert.Equal(string.Empty, labelled.Rows[ReviewLogSchema.CheckFindingTable].Single()[Index(ReviewLogSchema.CheckFindingColumns, "message")]);
            foreach (string cell in record.Rows.Values.SelectMany(rows => rows).SelectMany(r => r))
                TestAssert.False(cell.Contains("SYNTH-PAT-001"), "Clear patient ID leaked: " + cell);
        }

        public static void WriterAppendsAndGuardsHeaders()
        {
            string directory = TempDirectory();
            try
            {
                var writer = new ReviewLogWriter(directory);
                writer.Write(Build(SyntheticScenarioFactory.Create("baseline-pass"), Context(ReviewLogModes.Live)));
                writer.Write(Build(SyntheticScenarioFactory.Create("mixed-review"), Context(ReviewLogModes.Retrospective)));
                foreach (string table in ReviewLogSchema.WriteOrder)
                {
                    string text = File.ReadAllText(Path.Combine(directory, ReviewLogSchema.FileName(table)), new UTF8Encoding(false, true));
                    TestAssert.False(text.Length > 0 && text[0] == '\uFEFF', "No BOM.");
                    var rows = ReviewLogCsv.Parse(text);
                    TestAssert.Equal(string.Join(",", ReviewLogSchema.Columns(table)), string.Join(",", rows[0]));
                    TestAssert.Equal(1, rows.Count(r => r[0] == ReviewLogSchema.Columns(table)[0]), "Header written once: " + table);
                    TestAssert.True(rows.Skip(1).All(r => r.Count == ReviewLogSchema.Columns(table).Count));
                }
                TestAssert.Equal(3, ReviewLogCsv.Parse(File.ReadAllText(Path.Combine(directory, "review.csv"))).Count);

                string other = TempDirectory();
                try
                {
                    File.WriteAllText(Path.Combine(other, "review.csv"), "review_id,old_column\r\nx,y\r\n");
                    TestAssert.Throws<ReviewLogSchemaException>(() => new ReviewLogWriter(other).Write(Build(MinimalSnapshot(), Context(ReviewLogModes.Live))));
                    TestAssert.False(File.Exists(Path.Combine(other, "goal_result.csv")), "No partial rows after a schema mismatch.");
                    TestAssert.Equal("review_id,old_column\r\nx,y\r\n", File.ReadAllText(Path.Combine(other, "review.csv")));
                }
                finally { Directory.Delete(other, true); }
            }
            finally { Directory.Delete(directory, true); }
        }

        public static void ConcurrentAppendsAreSafe()
        {
            string directory = TempDirectory();
            try
            {
                var snapshot = SyntheticScenarioFactory.Create("mixed-review");
                int expectedChecks = 0;
                Parallel.For(0, 24, new ParallelOptions { MaxDegreeOfParallelism = 8 }, i =>
                {
                    var record = Build(snapshot, Context(i % 2 == 0 ? ReviewLogModes.Live : ReviewLogModes.Headless));
                    System.Threading.Interlocked.Add(ref expectedChecks, record.Rows[ReviewLogSchema.CheckFindingTable].Count);
                    new ReviewLogWriter(directory, 30000).Write(record);
                });
                var reviews = ReviewLogCsv.Parse(File.ReadAllText(Path.Combine(directory, "review.csv")));
                TestAssert.Equal(25, reviews.Count);
                TestAssert.Equal(24, reviews.Skip(1).Select(r => r[1]).Distinct().Count());
                var checks = ReviewLogCsv.Parse(File.ReadAllText(Path.Combine(directory, "check_finding.csv")));
                TestAssert.Equal(expectedChecks + 1, checks.Count);
                TestAssert.True(checks.Skip(1).All(r => r.Count == ReviewLogSchema.CheckFindingColumns.Length));
            }
            finally { Directory.Delete(directory, true); }
        }

        public static void RulesetResolverHashesAndPinsRevisions()
        {
            string directory = TempDirectory();
            try
            {
                string root = Path.Combine(directory, "Configuration");
                var store = new ClearPlan.Core.Configuration.ConfigurationHistoryStore(root);
                byte[] content = Encoding.UTF8.GetBytes("{\"synthetic\":true}");
                var revision = store.Import("default-rules", "default_rules", ".json", content, "synthetic-actor", "synthetic test", 0, null);
                string managed = store.GetManagedPath("default-rules");
                var described = ReviewLogRulesetResolver.Describe("default-review-rules", "default_rules", managed, root);
                TestAssert.Equal("available", described.Availability);
                TestAssert.Equal(revision.Sha256, described.Sha256);
                TestAssert.Equal(ReviewLogRulesetResolver.FileSha256(managed), described.Sha256);
                TestAssert.Equal(revision.RevisionNumber, described.RevisionNumber.Value);
                TestAssert.True(described.RevisionCreatedUtc.HasValue);
                string external = Path.Combine(directory, "external.json");
                File.Copy(managed, external);
                var plain = ReviewLogRulesetResolver.Describe("default-review-rules", "default_rules", external, root);
                TestAssert.Equal(described.Sha256, plain.Sha256);
                TestAssert.False(plain.RevisionNumber.HasValue);
                TestAssert.Equal("not_configured", ReviewLogRulesetResolver.Describe("x", "y", null, root).Availability);
                TestAssert.Equal("unavailable", ReviewLogRulesetResolver.Describe("x", "y", Path.Combine(directory, "missing.json"), root).Availability);
                string one = ReviewLogRecordBuilder.RulesetSha256(new[] { described }, "abc");
                string two = ReviewLogRecordBuilder.RulesetSha256(new[] { plain }, "abc");
                TestAssert.Equal(one, two);
                TestAssert.False(one == ReviewLogRecordBuilder.RulesetSha256(new[] { described }, "abd"), "Goal-table change must change the rule-set hash.");
            }
            finally { Directory.Delete(directory, true); }
        }

        public static void SettingsRoundTripReviewLogPaths()
        {
            var source = new ClearPlanSettingsModel();
            source.Paths.ReviewLogDirectory = @"\\server\share\ReviewLog";
            source.Paths.ReviewLogKeyFile = @"D:\Protected\review-log.key";
            var restored = new ClearPlanSettingsModel();
            PathSettingsIni.Apply(restored, PathSettingsIni.Serialize(source));
            TestAssert.Equal(source.Paths.ReviewLogDirectory, restored.Paths.ReviewLogDirectory);
            TestAssert.Equal(source.Paths.ReviewLogKeyFile, restored.Paths.ReviewLogKeyFile);
            var defaults = new ClearPlanSettingsModel();
            TestAssert.False(defaults.ReviewLog.Enabled, "Review logging is opt-in.");
            TestAssert.False(defaults.ReviewLog.LogClearIdentifiers, "Pseudonymisation is the default.");
            TestAssert.False(defaults.ReviewLog.LogPlanLabel);
            var json = Newtonsoft.Json.JsonConvert.DeserializeObject<ClearPlanSettingsModel>("{\"reviewLog\":{\"enabled\":true}}");
            TestAssert.True(json.ReviewLog.Enabled);
            TestAssert.False(json.ReviewLog.LogClearIdentifiers);
        }

        public static void NativeEntriesStayReadOnly()
        {
            string mining = File.ReadAllText(Path.Combine("tools", "review-log", "mining", "RetrospectiveReviewMining.cs"));
            string adapter = File.ReadAllText(Path.Combine("ClearPlan.Script", "Review", "EsapiReviewLog.cs"));
            string live = File.ReadAllText(Path.Combine("ClearPlan.Script", "MainView.ReviewLog.cs"));
            // The internal headless report tool is not part of the public source release; check it only where present.
            string headlessPath = Path.Combine("tools", "clinical-report", "HeadlessPlanReport.cs");
            string headless = File.Exists(headlessPath) ? File.ReadAllText(headlessPath) : null;
            string mainView = File.ReadAllText(Path.Combine("ClearPlan.Script", "MainView.xaml.cs"));
            foreach (string source in new[] { mining, adapter, live })
            {
                TestAssert.False(source.Contains("BeginModifications") || source.Contains("SaveModifications"), "Review-log paths must not modify the TPS.");
                TestAssert.False(source.Contains("ReportPdf") || source.Contains("Aria"), "Review-log paths write no report and upload nothing.");
            }
            TestAssert.True(mining.Contains("[assembly: ESAPIScript(IsWriteable = false)]"));
            TestAssert.True(mining.Contains("ReviewLogModes.Retrospective") && mining.Contains("app.ClosePatient()") && mining.Contains("OpenPatientById"));
            TestAssert.True(mining.Contains("new EsapiReviewSnapshotBuilder().Build(vm, settings)"), "Retrospective mode reuses the review snapshot builder.");
            if (headless != null)
                TestAssert.True(headless.Contains("ReviewLogModes.Headless") && headless.Contains("EsapiReviewLog.TryWrite"));
            TestAssert.True(live.Contains("ReviewLogModes.Live"));
            TestAssert.True(mainView.Contains("WriteLiveReviewLog(snapshot, dialog.FileName);") && mainView.Contains("WriteLiveReviewLogOnClose();"));
            TestAssert.True(mainView.Contains("private void LogLiveMining()") && mainView.Contains("\"LiveMining.csv\""), "Legacy LiveMining.csv writer stays.");
            TestAssert.True(adapter.Contains("catch (Exception error)") && adapter.Contains("maxWaitMilliseconds"), "Native logging is guarded and time-bounded.");
        }

        public static void ExampleExportMatchesExpectations()
        {
            string directory = TempDirectory();
            try
            {
                var expected = WriteExample(directory);
                foreach (var pair in expected.Where(p => p.Key.StartsWith("rows_", StringComparison.Ordinal)))
                {
                    string table = pair.Key.Substring(5);
                    int rows = ReviewLogCsv.Parse(File.ReadAllText(Path.Combine(directory, ReviewLogSchema.FileName(table)))).Count - 1;
                    TestAssert.Equal(pair.Value, rows, table);
                }
            }
            finally { Directory.Delete(directory, true); }
        }

        /// <summary>Program entry: ExportReviewLogExample &lt;dir&gt;. Synthetic snapshots only; no patient data.</summary>
        public static int ExportExample(string directory)
        {
            if (string.IsNullOrWhiteSpace(directory)) { Console.Error.WriteLine("ExportReviewLogExample <empty target directory>"); return 2; }
            if (Directory.Exists(directory) && Directory.EnumerateFileSystemEntries(directory).Any()) { Console.Error.WriteLine("Target directory must be empty."); return 2; }
            Directory.CreateDirectory(directory);
            var expected = WriteExample(directory);
            File.WriteAllText(Path.Combine(directory, "expected_counts.json"), "{\n" + string.Join(",\n",
                expected.OrderBy(p => p.Key, StringComparer.Ordinal).Select(p => "  \"" + p.Key + "\": " + p.Value.ToString(CultureInfo.InvariantCulture))) + "\n}\n");
            Console.WriteLine("Synthetic review-log example written: " + expected["rows_review"] + " review rows.");
            return 0;
        }

        private static Dictionary<string, int> WriteExample(string directory)
        {
            var writer = new ReviewLogWriter(directory);
            var pseudonymizer = new ReviewLogPseudonymizer(SyntheticKey, false);
            var options = new ReviewLogOptions { LogMessages = true };
            string[] modes = { ReviewLogModes.Retrospective, ReviewLogModes.Retrospective, ReviewLogModes.Live, ReviewLogModes.Headless,
                ReviewLogModes.Live, ReviewLogModes.Live, ReviewLogModes.Headless };
            var counts = new Dictionary<string, int>(StringComparer.Ordinal);
            Action<string, int> add = (k, v) => { int c; counts.TryGetValue(k, out c); counts[k] = c + v; };
            IList<string> ids = SyntheticScenarioFactory.ScenarioIds;
            for (int i = 0; i < ids.Count; i++)
            {
                ReviewSnapshot snapshot = SyntheticScenarioFactory.Create(ids[i]);
                var started = new DateTimeOffset(2026, 7 + i / 3, 3 + 4 * (i % 3), 8 + i, 15, 0, TimeSpan.Zero);
                var context = new ReviewLogContext
                {
                    ReviewId = new Guid(i + 1, 0, 0, new byte[] { 0, 0, 0, 0, 0, 0, 0, (byte)(i + 1) }),
                    Mode = modes[i], StartedUtc = started, FinishedUtc = started.AddMinutes(6 + i),
                    ClearPlanVersion = "3.2.0-synthetic", BuildId = "synthetic-build",
                    UserId = "SYNTH\\physicist" + (i % 2 + 1), PatientId = "SYNTH-PAT-" + (i % 4 + 1).ToString("000", CultureInfo.InvariantCulture),
                    CourseId = "C1", PlanUid = "1.2.826.0.1.3680043.10.9999." + (i + 1), PlanId = "SYNTH_" + ids[i], PlanningItemType = "plan",
                    PlanCreatedUtc = started.AddDays(-2), PlanApprovalStatus = i < 2 ? "TreatmentApproved" : "PlanningApproved",
                    PlanningApprovalUtc = started.AddDays(-1), TreatmentApprovalUtc = i < 2 ? started.AddHours(-3) : (DateTimeOffset?)null,
                    GoalTableId = "Synthetic H&N 30x2Gy", ReportWritten = modes[i] != ReviewLogModes.Retrospective,
                    ReportSha256 = modes[i] != ReviewLogModes.Retrospective ? ReviewLogRulesetResolver.Sha256Hex("synthetic report " + i) : null
                };
                // Two rule-set epochs: the default-rules document changes after the third review.
                context.ConfigurationInputs.Add(new ReviewLogConfigurationInput { ConfigKey = "constraints-excel", Category = "constraint_catalog",
                    Sha256 = ReviewLogRulesetResolver.Sha256Hex("synthetic workbook v1"), Availability = "available" });
                context.ConfigurationInputs.Add(new ReviewLogConfigurationInput { ConfigKey = "default-review-rules", Category = "default_rules",
                    RevisionNumber = i < 3 ? 1 : 2, RevisionCreatedUtc = i < 3 ? new DateTimeOffset(2026, 6, 1, 0, 0, 0, TimeSpan.Zero) : new DateTimeOffset(2026, 8, 1, 0, 0, 0, TimeSpan.Zero),
                    Sha256 = ReviewLogRulesetResolver.Sha256Hex(i < 3 ? "synthetic defaults r1" : "synthetic defaults r2"), Availability = "available" });
                context.ConfigurationInputs.Add(new ReviewLogConfigurationInput { ConfigKey = "plancheck-selection", Category = "check_selection", Availability = "not_configured" });
                ReviewLogRecord record = ReviewLogRecordBuilder.Build(snapshot, context, pseudonymizer, options);
                writer.Write(record);
                foreach (string table in ReviewLogSchema.WriteOrder) add("rows_" + table, record.Rows[table].Count);
                add("reviews_" + modes[i], 1);
                foreach (ReviewPqmRow row in snapshot.PqmRows) add("goal_" + ReviewLogGoalStatus.FromSnapshot(row.Status), 1);
                foreach (ReviewCheckRow row in snapshot.PlanCheckRows) add("check_" + ReviewLogCheckStatus.FromSnapshot(row.Status), 1);
                add("parameters_available", record.Rows[ReviewLogSchema.PlanParameterTable].Count(r => r[Index(ReviewLogSchema.PlanParameterColumns, "availability")] == "available"));
            }
            // One failed retrospective lookup: identifiers only, no evaluated content.
            var failed = new ReviewLogContext
            {
                ReviewId = new Guid(99, 0, 0, new byte[] { 0, 0, 0, 0, 0, 0, 0, 99 }), Mode = ReviewLogModes.Retrospective,
                StartedUtc = new DateTimeOffset(2026, 9, 20, 7, 0, 0, TimeSpan.Zero), FinishedUtc = new DateTimeOffset(2026, 9, 20, 7, 0, 2, TimeSpan.Zero),
                ClearPlanVersion = "3.2.0-synthetic", BuildId = "synthetic-build", UserId = "SYNTH\\physicist1", PatientId = "SYNTH-PAT-404",
                CourseId = "C1", PlanId = "SYNTH_missing", PlanningItemType = "plan", RunStatus = ReviewLogRunStatus.Failed, RunMessage = "plan-not-found"
            };
            ReviewLogRecord failedRecord = ReviewLogRecordBuilder.Build(null, failed, pseudonymizer, options);
            writer.Write(failedRecord);
            foreach (string table in ReviewLogSchema.WriteOrder) add("rows_" + table, failedRecord.Rows[table].Count);
            add("reviews_retrospective", 1);
            add("reviews_failed", 1);
            add("ruleset_epochs", 2);
            return counts;
        }

        private static ReviewSnapshot MinimalSnapshot()
        {
            var snapshot = new ReviewSnapshot { Synthetic = true, ActivePlanKey = "p" };
            snapshot.Plans.Add(new ReviewPlanRow { PlanKey = "p", TotalDoseGy = 60, DosePerFractionGy = 2, FractionCount = 30 });
            return snapshot;
        }

        private static ReviewLogContext Context(string mode)
        {
            return new ReviewLogContext { Mode = mode, StartedUtc = DateTimeOffset.UtcNow, UserId = "SYNTH\\user", PatientId = "SYNTH-PAT-001",
                CourseId = "C1", PlanUid = "1.2.3.4", PlanId = "SYNTH_PLAN", GoalTableId = "Synthetic table" };
        }

        private static ReviewLogRecord Build(ReviewSnapshot snapshot, ReviewLogContext context)
        {
            return ReviewLogRecordBuilder.Build(snapshot, context, new ReviewLogPseudonymizer(SyntheticKey, false), new ReviewLogOptions());
        }

        private static int Index(IList<string> columns, string name)
        {
            int index = columns.IndexOf(name);
            if (index < 0) throw new InvalidOperationException("Unknown column " + name);
            return index;
        }

        private static string TempDirectory()
        {
            string path = Path.Combine(Path.GetTempPath(), "clearplan-review-log-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(path);
            return path;
        }
    }
}
