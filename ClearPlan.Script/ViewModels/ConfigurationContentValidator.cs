using System;
using System.IO;
using System.Linq;
using ClearPlan.Core.Configuration;
using ClearPlan.Core.Constraints;
using ClearPlan.Core.Fields;
using ClearPlan.Core.PlanAnalysis;
using ClearPlan.Core.Review;

namespace ClearPlan
{
    public static class ConfigurationContentValidator
    {
        public static void Validate(string key, byte[] content, string managedRoot)
        {
            if (content == null || content.Length == 0 || content.Length > 32 * 1024 * 1024)
                throw new FormatException("Konfigurationsdatei ist leer oder zu groß.");
            if (key == "constraints" || key == "refdb")
            {
                // The existing workbook/catalog readers require a path. Only a private validation
                // file beneath the explicitly configured managed root is written; no source is opened for writing.
                var store = new ConfigurationHistoryStore(managedRoot);
                string path = store.CreateTemporaryFile(".validation", key == "constraints" ? ".xlsx" : ".json", content);
                try
                {
                    ConstraintCatalog catalog = key == "constraints"
                        ? new ExcelConstraintSource(true).Load(path) : new RefDbJsonConstraintSource(true).Load(path);
                    var issues = catalog.Issues.Concat(ConstraintCatalogValidator.Validate(catalog)).Where(issue => issue.IsFatal).ToList();
                    if (issues.Count > 0 || catalog.Tables.Count == 0)
                        throw new FormatException("Constraint-Katalog ist ungültig oder enthält keine Tabellen. " +
                            string.Join(" ", issues.Take(4).Select(issue => issue.Message)));
                }
                finally { store.DeleteTemporaryFile(path); }
                return;
            }
            string json = ConfigurationWorkspaceViewModel.Decode(content);
            if (key == "default-rules")
            {
                if (content.Length > 131072) throw new FormatException("Default-Regeln dürfen 128 KiB nicht überschreiten.");
                var parsed = DefaultReviewRuleConfiguration.Parse(json);
                if (parsed.Status == ReviewStatusCodes.Unavailable) throw new FormatException(parsed.Message);
            }
            else if (key == "field-naming")
            {
                if (content.Length > 32768) throw new FormatException("Feldnamen-Regeln dürfen 32 KiB nicht überschreiten.");
                var parsed = FieldNamingRuleConfiguration.Parse(json);
                if (parsed.Status == ReviewStatusCodes.Unavailable) throw new FormatException(parsed.Message);
            }
            else if (key == "plancheck-selection")
            {
                if (content.Length > 131072) throw new FormatException("PlanCheck-Auswahl darf 128 KiB nicht überschreiten.");
                var parsed = PlanCheckSelectionConfiguration.Parse(json);
                if (parsed.Status == ReviewStatusCodes.Unavailable) throw new FormatException(parsed.Message);
            }
            else if (key == "aliases") StructureAliasConfiguration.Deserialize(json);
            else if (key == "aria-upload") ClearPlan.Core.Integration.AriaUploadConfiguration.Parse(json);
            else if (key == "mlc-profiles") NativeMlcProfileCatalog.Parse(json);
            else if (key == "dose-rate-profiles")
            {
                if (content.Length > 65536) throw new FormatException("Dosisraten-Profile dürfen 64 KiB nicht überschreiten.");
                DoseRateEstimationProfileCatalog.Parse(json);
            }
            else throw new FormatException("Unbekannte Konfigurationsart; keine Übernahme möglich.");
        }
    }
}
