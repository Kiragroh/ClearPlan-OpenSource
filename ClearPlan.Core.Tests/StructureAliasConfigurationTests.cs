using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClearPlan.Core.Constraints;
using ClearPlan.Core.Settings;

namespace ClearPlan.Core.Tests
{
    internal static class StructureAliasConfigurationTests
    {
        public static void RejectsAmbiguousAndCrossSideAliases()
        {
            var rows = new[]
            {
                new StructureDefinition { CanonicalName = "Kidney_L", Aliases = { "Renal" } },
                new StructureDefinition { CanonicalName = "Kidney_R", Aliases = { "Re nal" } }
            };
            var issues = StructureAliasConfiguration.Validate(rows);
            TestAssert.True(issues.Any(issue => issue.Contains("Renal") || issue.Contains("renal")));
            rows[1].Aliases.Clear();
            rows[0].SideAliasesLeft.Add("Kidney_Left");
            rows[0].SideAliasesRight.Add("Kidney Left");
            TestAssert.True(StructureAliasConfiguration.Validate(rows).Count > 0);
        }

        public static void PreservesSourceIdentityAndClinicalValues()
        {
            var source = new[]
            {
                new StructureDefinition { StructureId = "local-id", CanonicalName = "Cord", Active = true,
                    Aliases = { "ExistingAlias" }, Codes = { "CODE" } }
            };
            var configuration = new[]
            {
                new StructureDefinition { CanonicalName = "SpinalCord", Aliases = { "Cord", "Myelon" }, Active = true }
            };
            var merged = StructureAliasConfiguration.Merge(source, configuration);
            TestAssert.Equal("local-id", merged.Single().StructureId);
            TestAssert.Equal("SpinalCord", merged.Single().CanonicalName);
            TestAssert.False(merged.Single().Aliases.Contains("ExistingAlias"), "Explicit overrides can remove a source alias without editing the source file.");
            TestAssert.True(merged.Single().Codes.Contains("CODE"));
            TestAssert.Equal("Cord", source.Single().CanonicalName, "The source must not be mutated.");
        }

        public static void RoundTripsJsonAndImportsExistingSources()
        {
            var rows = new[] { new StructureDefinition { CanonicalName = "SpinalCord", Active = true, Aliases = { "Rückenmark" } } };
            string json = StructureAliasConfiguration.Serialize(rows);
            TestAssert.True(json.Contains("canonical_name"));
            var loaded = StructureAliasConfiguration.Deserialize(json);
            TestAssert.Equal("Rückenmark", loaded.Single().Aliases.Single());
            string fixtures = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Fixtures");
            var excel = StructureAliasConfiguration.Import(Path.Combine(fixtures, "excel-catalog-fixture.xlsx"));
            var refDb = StructureAliasConfiguration.Import(Path.Combine(fixtures, "refdb-minimal.json"));
            TestAssert.True(excel.Count > 0);
            TestAssert.True(refDb.Any(row => row.CanonicalName == "SpinalCord"));
            TestAssert.Throws<FormatException>(() => StructureAliasConfiguration.Deserialize("{\"schema\":\"unknown\",\"structures\":[]}"));
        }

        public static void PersistsOptionalPathInIni()
        {
            var settings = new ClearPlanSettingsModel();
            settings.ConstraintSource.StructureAliasesJsonPath = "Mappings/structure-aliases.json";
            string ini = PathSettingsIni.Serialize(settings);
            var reloaded = new ClearPlanSettingsModel();
            PathSettingsIni.Apply(reloaded, ini);
            TestAssert.Equal("Mappings/structure-aliases.json", reloaded.ConstraintSource.StructureAliasesJsonPath);
            TestAssert.Equal(ConstraintSourceMode.Automatic, reloaded.ConstraintSource.Mode);
        }

        public static void PreservesTg263SemanticMarkers()
        {
            TestAssert.False(StructureAliasResolver.NormalizeName("Lungs~") == StructureAliasResolver.NormalizeName("Lungs"),
                "A partial organ must not normalize to a whole organ.");
            TestAssert.False(StructureAliasResolver.NormalizeName("PTV-03") == StructureAliasResolver.NormalizeName("PTV03"),
                "A cropped target must not normalize to a numbered target.");
        }

        public static void RejectsUnknownLaterality()
        {
            TestAssert.True(StructureAliasConfiguration.Validate(new[]
            {
                new StructureDefinition { CanonicalName = "Kidney_L", Laterality = "typo" }
            }).Count > 0);
        }

        public static void BlocksMissingOverridesWithoutChangingSource()
        {
            string fixtures = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Fixtures");
            var options = new ConstraintSourceOptions
            {
                Mode = ConstraintSourceMode.Automatic, RefDbJsonPath = "refdb-minimal.json",
                ExcelWorkbookPath = "excel-catalog-fixture.xlsx", StructureAliasesJsonPath = "missing-alias.json"
            };
            var loaded = new ConstraintCatalogService().Load(options, fixtures);
            TestAssert.False(loaded.IsUsable);
            TestAssert.Equal("RefDB", loaded.ActiveSource, "Alias failure must not switch clinical constraint source.");
            TestAssert.True(loaded.Errors.Any(error => error.Contains("Aliase")));
        }

        public static void SavesOnlyAliasDocumentsAndKeepsBackup()
        {
            string directory = Path.Combine(Path.GetTempPath(), "ClearPlanAliasTests-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            try
            {
                string path = Path.Combine(directory, "aliases.json");
                var rows = new[] { new StructureDefinition { CanonicalName = "SpinalCord", Aliases = { "Cord" } } };
                StructureAliasConfiguration.Save(path, rows);
                string original = File.ReadAllText(path);
                rows[0].Aliases.Add("Myelon");
                StructureAliasConfiguration.Save(path, rows);
                TestAssert.Equal(original, File.ReadAllText(path + ".bak"));
                TestAssert.True(StructureAliasConfiguration.Import(path).Single().Aliases.Contains("Myelon"));
                string source = Path.Combine(directory, "source.json");
                File.Copy(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Fixtures", "refdb-minimal.json"), source);
                string originalSource = File.ReadAllText(source);
                TestAssert.Throws<FormatException>(() => StructureAliasConfiguration.Save(source, rows));
                TestAssert.Equal(originalSource, File.ReadAllText(source));
            }
            finally { Directory.Delete(directory, true); }
        }

        public static void PublicExampleKeepsThresholdsAndSeparatesOrgans()
        {
            string fixtures = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Fixtures");
            var options = new ConstraintSourceOptions { Mode = ConstraintSourceMode.Excel,
                ExcelWorkbookPath = "public-starter-constraints.xlsx" };
            var baseline = new ConstraintCatalogService().Load(options, fixtures);
            TestAssert.True(baseline.IsUsable);
            options.StructureAliasesJsonPath = "tg263-aliases-example.json";
            var configured = new ConstraintCatalogService().Load(options, fixtures);
            TestAssert.True(configured.IsUsable, string.Join("; ", configured.Errors));
            var before = baseline.Catalog.Tables.SelectMany(table => table.Constraints).ToList();
            var after = configured.Catalog.Tables.SelectMany(table => table.Constraints).ToList();
            TestAssert.Equal(before.Count, after.Count);
            for (int index = 0; index < before.Count; index++)
            {
                TestAssert.Equal(before[index].Goal, after[index].Goal);
                TestAssert.Equal(before[index].Variation, after[index].Variation);
                TestAssert.Equal(before[index].Unit, after[index].Unit);
                TestAssert.Equal(before[index].TableId, after[index].TableId);
            }
            StructureDefinition cord = configured.Catalog.Structures.Single(row => row.CanonicalName == "SpinalCord");
            TestAssert.False(cord.Aliases.Contains("Spinalkanal"));
            TestAssert.True(configured.Catalog.Structures.Any(row => row.CanonicalName == "SpinalCanal"));
            StructureDefinition lungs = configured.Catalog.Structures.Single(row => row.CanonicalName == "Lungs");
            TestAssert.False(lungs.Aliases.Contains("Lung L"));
            TestAssert.False(lungs.Aliases.Contains("Lung R"));
        }
    }
}
