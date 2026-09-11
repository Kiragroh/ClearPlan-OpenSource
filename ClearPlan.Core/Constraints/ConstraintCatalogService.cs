using System;
using System.Collections.Generic;
using System.Linq;
using ClearPlan.Core.Settings;

namespace ClearPlan.Core.Constraints
{
    public sealed class ConstraintCatalogService
    {
        public ConstraintCatalogLoadResult Load(
            ConstraintSourceOptions options,
            string baseDirectory)
        {
            ConstraintCatalogLoadResult result = LoadSource(options, baseDirectory);
            ApplyAliases(result, options == null ? null : options.StructureAliasesJsonPath, baseDirectory);
            return result;
        }

        public static void ApplyAliases(ConstraintCatalogLoadResult result, string configuredPath, string baseDirectory)
        {
            if (result == null || !result.IsUsable || string.IsNullOrWhiteSpace(configuredPath)) return;
            try
            {
                string path = SettingsPathResolver.Resolve(baseDirectory, configuredPath);
                // A runtime override must be an explicit alias document, never another constraint catalog.
                IList<StructureDefinition> aliases = StructureAliasConfiguration.Deserialize(
                    System.IO.File.ReadAllText(path, System.Text.Encoding.UTF8));
                result.Catalog.Structures = StructureAliasConfiguration.Merge(result.Catalog.Structures, aliases);
                result.Warnings.Add("Externe Struktur-Aliase aktiv: " + path);
            }
            catch (Exception exception)
            {
                result.Errors.Add("Struktur-Aliase nicht angewendet: " + exception.Message);
            }
        }

        private static ConstraintCatalogLoadResult LoadSource(
            ConstraintSourceOptions options,
            string baseDirectory)
        {
            var result = new ConstraintCatalogLoadResult();
            if (options == null)
            {
                result.Errors.Add("Constraint source settings are missing.");
                return result;
            }

            switch (options.Mode)
            {
                case ConstraintSourceMode.RefDb:
                    return LoadExplicit(
                        "RefDB",
                        options.RefDbJsonPath,
                        baseDirectory,
                        path => new RefDbJsonConstraintSource(
                            options.IncludeInactiveTables).Load(path));

                case ConstraintSourceMode.Excel:
                    return LoadExplicit(
                        "Excel",
                        options.ExcelWorkbookPath,
                        baseDirectory,
                        path => new ExcelConstraintSource(
                            options.IncludeInactiveTables).Load(path));

                case ConstraintSourceMode.Automatic:
                default:
                    return LoadAutomatic(options, baseDirectory);
            }
        }

        private static ConstraintCatalogLoadResult LoadAutomatic(
            ConstraintSourceOptions options,
            string baseDirectory)
        {
            var result = new ConstraintCatalogLoadResult();
            ConstraintCatalog refDbCatalog = TryLoad(
                options.RefDbJsonPath,
                baseDirectory,
                path => new RefDbJsonConstraintSource(
                    options.IncludeInactiveTables).Load(path));
            if (IsCatalogUsable(refDbCatalog))
            {
                Activate(result, refDbCatalog);
                return result;
            }

            AddAttemptWarnings(result.Warnings, "RefDB", refDbCatalog);
            ConstraintCatalog excelCatalog = TryLoad(
                options.ExcelWorkbookPath,
                baseDirectory,
                path => new ExcelConstraintSource(
                    options.IncludeInactiveTables).Load(path));
            if (IsCatalogUsable(excelCatalog))
            {
                Activate(result, excelCatalog);
                result.Warnings.Add("RefDB nicht verfügbar – Excel aktiv.");
                return result;
            }

            AddAttemptWarnings(result.Warnings, "Excel", excelCatalog);
            result.Catalog = excelCatalog ?? refDbCatalog;
            result.Errors.Add("Neither RefDB nor Excel provides a usable constraint catalog.");
            return result;
        }

        private static ConstraintCatalogLoadResult LoadExplicit(
            string sourceName,
            string configuredPath,
            string baseDirectory,
            Func<string, ConstraintCatalog> loader)
        {
            var result = new ConstraintCatalogLoadResult();
            ConstraintCatalog catalog = TryLoad(configuredPath, baseDirectory, loader);
            result.Catalog = catalog;
            if (IsCatalogUsable(catalog))
            {
                Activate(result, catalog);
                return result;
            }

            foreach (string error in AttemptMessages(sourceName, catalog))
            {
                result.Errors.Add(error);
            }

            if (result.Errors.Count == 0)
            {
                result.Errors.Add(sourceName + " did not provide a usable constraint catalog.");
            }

            return result;
        }

        private static ConstraintCatalog TryLoad(
            string configuredPath,
            string baseDirectory,
            Func<string, ConstraintCatalog> loader)
        {
            string path = SettingsPathResolver.Resolve(baseDirectory, configuredPath);
            return loader(path);
        }

        private static bool IsCatalogUsable(ConstraintCatalog catalog)
        {
            return catalog != null &&
                   catalog.Tables.Count > 0 &&
                   !catalog.Issues.Any(issue => issue.IsFatal);
        }

        private static void Activate(
            ConstraintCatalogLoadResult result,
            ConstraintCatalog catalog)
        {
            result.Catalog = catalog;
            result.ActiveSource = catalog.SourceKind ?? string.Empty;
            result.ActivePath = catalog.SourcePath ?? string.Empty;
            foreach (CatalogValidationIssue warning in catalog.Issues
                .Where(issue => !issue.IsFatal))
            {
                result.Warnings.Add(warning.Message);
            }
        }

        private static void AddAttemptWarnings(
            ICollection<string> warnings,
            string sourceName,
            ConstraintCatalog catalog)
        {
            foreach (string message in AttemptMessages(sourceName, catalog))
            {
                warnings.Add(message);
            }
        }

        private static IEnumerable<string> AttemptMessages(
            string sourceName,
            ConstraintCatalog catalog)
        {
            if (catalog == null)
            {
                yield return sourceName + ": source could not be loaded.";
                yield break;
            }

            foreach (CatalogValidationIssue issue in catalog.Issues.Where(item => item.IsFatal))
            {
                yield return sourceName + ": " + issue.Message;
            }

            if (catalog.Tables.Count == 0)
            {
                yield return sourceName + ": no active constraint tables found.";
            }
        }
    }
}
