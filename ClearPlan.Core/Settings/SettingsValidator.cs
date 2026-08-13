using System;
using System.Collections.Generic;
using System.IO;
using ClearPlan.Core.Constraints;

namespace ClearPlan.Core.Settings
{
    public static class SettingsValidator
    {
        public static SettingsValidationResult Validate(
            ClearPlanSettingsModel settings,
            string baseDirectory)
        {
            var result = new SettingsValidationResult();
            if (settings == null)
            {
                result.Errors.Add("Einstellungen fehlen.");
                return result;
            }

            ValidateSources(settings.ConstraintSource, baseDirectory, result);
            ClearPlanPathOptions paths = settings.Paths ?? new ClearPlanPathOptions();
            ValidateWritableDirectory(
                "Logs",
                paths.LogsDirectory,
                baseDirectory,
                result);
            ValidateWritableDirectory(
                "Reports",
                paths.ReportsDirectory,
                baseDirectory,
                result);
            ValidateWritableDirectory(
                "Exports",
                paths.CsvExportDirectory,
                baseDirectory,
                result);
            ValidateWritableDirectory(
                "State",
                paths.StateDirectory,
                baseDirectory,
                result);
            ValidateWritableDirectory(
                "Settings",
                baseDirectory,
                baseDirectory,
                result);
            return result;
        }

        private static void ValidateSources(
            ConstraintSourceOptions options,
            string baseDirectory,
            SettingsValidationResult result)
        {
            options = options ?? new ConstraintSourceOptions();
            bool refDbValid = ValidateOptionalFile(
                "RefDB",
                options.RefDbJsonPath,
                ".json",
                baseDirectory,
                options.Mode == ConstraintSourceMode.RefDb,
                result);
            bool excelValid = ValidateOptionalFile(
                "Excel",
                options.ExcelWorkbookPath,
                ".xlsx",
                baseDirectory,
                options.Mode == ConstraintSourceMode.Excel,
                result);
            if (options.Mode == ConstraintSourceMode.Automatic &&
                !refDbValid &&
                !excelValid)
            {
                result.Errors.Add(
                    "Automatic benötigt mindestens eine lesbare RefDB-JSON- oder Excel-Datei.");
            }
            else if (options.Mode == ConstraintSourceMode.Automatic &&
                     !refDbValid &&
                     excelValid &&
                     !string.IsNullOrWhiteSpace(options.RefDbJsonPath))
            {
                result.Warnings.Add(
                    "RefDB nicht verfügbar – Excel kann als Fallback verwendet werden.");
            }
        }

        private static bool ValidateOptionalFile(
            string label,
            string path,
            string expectedExtension,
            string baseDirectory,
            bool required,
            SettingsValidationResult result)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                if (required)
                {
                    result.Errors.Add(label + "-Pfad ist erforderlich.");
                }

                return false;
            }

            string resolved;
            try
            {
                resolved = SettingsPathResolver.Resolve(baseDirectory, path);
            }
            catch (Exception exception)
            {
                result.Errors.Add(label + "-Pfad ist ungültig: " + exception.Message);
                return false;
            }

            if (!string.Equals(
                Path.GetExtension(resolved),
                expectedExtension,
                StringComparison.OrdinalIgnoreCase))
            {
                result.Errors.Add(
                    label + "-Datei muss die Endung " + expectedExtension + " haben.");
                return false;
            }

            if (!File.Exists(resolved))
            {
                if (required)
                {
                    result.Errors.Add(label + "-Datei wurde nicht gefunden.");
                }

                return false;
            }

            try
            {
                using (FileStream stream = File.Open(
                    resolved,
                    FileMode.Open,
                    FileAccess.Read,
                    FileShare.ReadWrite))
                {
                    return stream.CanRead;
                }
            }
            catch
            {
                result.Errors.Add(label + "-Datei ist nicht lesbar.");
                return false;
            }
        }

        private static void ValidateWritableDirectory(
            string label,
            string path,
            string baseDirectory,
            SettingsValidationResult result)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                result.Errors.Add(label + "-Verzeichnis fehlt.");
                return;
            }

            string resolved;
            try
            {
                resolved = SettingsPathResolver.Resolve(baseDirectory, path);
                Directory.CreateDirectory(resolved);
                string probe = Path.Combine(
                    resolved,
                    ".clearplan-write-" + Guid.NewGuid().ToString("N") + ".tmp");
                try
                {
                    File.WriteAllText(probe, string.Empty);
                }
                finally
                {
                    if (File.Exists(probe))
                    {
                        File.Delete(probe);
                    }
                }
            }
            catch
            {
                result.Errors.Add(label + "-Verzeichnis ist nicht beschreibbar.");
            }
        }
    }
}
