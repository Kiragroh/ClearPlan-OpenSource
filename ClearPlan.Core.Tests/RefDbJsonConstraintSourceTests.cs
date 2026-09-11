using System;
using System.IO;
using System.Linq;
using ClearPlan.Core.Constraints;

namespace ClearPlan.Core.Tests
{
    internal static class RefDbJsonConstraintSourceTests
    {
        public static void SummaryRangesSurviveSparseDetails()
        {
            string path = Path.Combine(Path.GetTempPath(), "clearplan-refdb-merge-" + Guid.NewGuid().ToString("N") + ".json");
            try
            {
                File.WriteAllText(path, @"{""schema"":""RSAlign.local_refdb_constraints.v1"",""structures"":[],""tables"":[{""id"":1,""name"":""Synthetic prescription"",""status"":""active"",""fx_min"":18,""fx_max"":20,""dpf_min"":2.5,""td_min"":45,""prescriptions"":[{""name"":""Synthetic Rx"",""status"":""active""}]}],""details"":{""1"":{""id"":1,""name"":""Synthetic prescription"",""fx_min"":null,""fx_max"":"""",""dpf_min"":null,""td_min"":null,""constraints"":[]}}}");
                var table = new RefDbJsonConstraintSource().Load(path).Tables.Single();
                TestAssert.Equal((int?)18, table.FractionCountMinimum);
                TestAssert.Equal((int?)20, table.FractionCountMaximum);
                TestAssert.Equal((decimal?)2.5m, table.DosePerFractionMinimumGy);
                TestAssert.Equal((decimal?)45m, table.TotalDoseMinimumGy);
                TestAssert.True(table.PrescriptionLabels.Contains("Synthetic Rx"));
            }
            finally { if (File.Exists(path)) File.Delete(path); }
        }

        public static void RejectsUnexpectedSchema()
        {
            string fixture = GetFixturePath();
            string temporaryPath = Path.Combine(
                Path.GetTempPath(),
                "clearplan-refdb-wrong-schema-" + Guid.NewGuid().ToString("N") + ".json");
            try
            {
                File.WriteAllText(
                    temporaryPath,
                    File.ReadAllText(fixture).Replace(
                        "RSAlign.local_refdb_constraints.v1",
                        "unsupported.schema"));

                ConstraintCatalog catalog = new RefDbJsonConstraintSource().Load(temporaryPath);

                TestAssert.Equal(0, catalog.Tables.Count);
                TestAssert.True(catalog.Issues.Any(issue => issue.IsFatal &&
                                                            issue.Field == "schema"));
            }
            finally
            {
                if (File.Exists(temporaryPath))
                {
                    File.Delete(temporaryPath);
                }
            }
        }

        public static void LoadsOnlyActiveEntitiesByDefault()
        {
            ConstraintCatalog catalog = new RefDbJsonConstraintSource().Load(GetFixturePath());

            TestAssert.Equal("RefDB", catalog.SourceKind);
            TestAssert.Equal(1, catalog.Tables.Count);
            TestAssert.Equal("201", catalog.Tables[0].TableId);
            TestAssert.Equal("table-active", catalog.Tables[0].DisplayName);
            TestAssert.Equal(2, catalog.Structures.Count);
            TestAssert.Equal(2, catalog.Statistics.TotalTables);
            TestAssert.Equal(1, catalog.Statistics.ActiveTables);
            TestAssert.Equal(3, catalog.Statistics.TotalStructures);
            TestAssert.Equal(2, catalog.Statistics.ActiveStructures);
        }

        public static void RetainsAliasesAndSideAliases()
        {
            ConstraintCatalog catalog = new RefDbJsonConstraintSource().Load(GetFixturePath());
            StructureDefinition kidney = catalog.Structures.Single(
                structure => structure.CanonicalName == "Kidney_L/R");

            TestAssert.True(kidney.Aliases.Contains("Kidney"));
            TestAssert.True(kidney.SideAliasesLeft.Contains("Niere_L"));
            TestAssert.True(kidney.SideAliasesRight.Contains("Niere_R"));
            TestAssert.Equal("L/R", kidney.Laterality);
        }

        public static void JoinsAndNormalizesConstraints()
        {
            ConstraintCatalog catalog = new RefDbJsonConstraintSource().Load(GetFixturePath());
            ConstraintTableDefinition table = catalog.Tables.Single();

            TestAssert.Equal(2, table.Constraints.Count);
            ConstraintDefinition dose = table.Constraints.Single(item => item.ConstraintId == "501");
            TestAssert.Equal("<=", dose.Comparator);
            TestAssert.Equal("cc", dose.Unit);
            TestAssert.Equal(18.5m, dose.Goal.Value);
            TestAssert.Equal(null, dose.Variation);
            TestAssert.Equal(2, dose.Priority.Value);

            ConstraintDefinition unresolved = table.Constraints.Single(item => item.ConstraintId == "502");
            TestAssert.Equal(null, unresolved.StructureId);
            TestAssert.Equal("Invented Liver", unresolved.RawStructureName);
            TestAssert.Equal(">=", unresolved.Comparator);
            TestAssert.Equal("CV21.5Gy", unresolved.Metric);
            TestAssert.False(catalog.Issues.Any(issue => issue.IsFatal));
        }

        private static string GetFixturePath()
        {
            return Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                "Fixtures",
                "refdb-minimal.json");
        }
    }
}
