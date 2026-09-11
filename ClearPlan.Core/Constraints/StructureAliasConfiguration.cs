using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ClearPlan.Core.Constraints
{
    /// <summary>Explicit local name mappings only; never modifies a constraint or TPS structure.</summary>
    public static class StructureAliasConfiguration
    {
        public const string Schema = "ClearPlan.structure_aliases.v1";
        public const string NomenclatureSource = "https://www.aapm.org/pubs/reports/RPT_263_Supplemental/default.asp";

        public static IList<string> Validate(IEnumerable<StructureDefinition> definitions)
        {
            var errors = new List<string>();
            var owners = new Dictionary<string, string>(StringComparer.Ordinal);
            var ids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            int rowNumber = 0;
            foreach (StructureDefinition row in definitions ?? Enumerable.Empty<StructureDefinition>())
            {
                rowNumber++;
                string location = "Zeile " + rowNumber;
                if (row == null || string.IsNullOrWhiteSpace(row.CanonicalName))
                {
                    errors.Add(location + ": kanonischer Name fehlt.");
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(row.StructureId) && !ids.Add(row.StructureId.Trim()))
                {
                    errors.Add(location + ": Struktur-ID ist doppelt: " + row.StructureId);
                }
                string laterality = (row.Laterality ?? string.Empty).Trim();
                if (!new[] { "", "L", "R", "L/R", "R+L" }.Contains(laterality, StringComparer.OrdinalIgnoreCase))
                {
                    errors.Add(location + ": Seite muss leer, L, R, L/R oder R+L sein.");
                }
                var left = new HashSet<string>(Names(row.SideAliasesLeft).Select(StructureAliasResolver.NormalizeName));
                var right = new HashSet<string>(Names(row.SideAliasesRight).Select(StructureAliasResolver.NormalizeName));
                if (left.Overlaps(right))
                {
                    errors.Add(location + " (" + row.CanonicalName + "): linker und rechter Alias überschneiden sich.");
                }
                string owner = location + " (" + row.CanonicalName + ")";
                foreach (string name in AllNames(row))
                {
                    string normalized = StructureAliasResolver.NormalizeName(name);
                    if (normalized.Length == 0 || name.Any(char.IsControl))
                    {
                        errors.Add(location + ": ungültiger Struktur- oder Aliasname.");
                        continue;
                    }
                    string existing;
                    if (owners.TryGetValue(normalized, out existing) && existing != owner)
                    {
                        errors.Add("Alias '" + name + "' (" + normalized + ") ist mehrdeutig: " + existing + " / " + owner + ".");
                    }
                    else
                    {
                        owners[normalized] = owner;
                    }
                }
            }
            return errors.Distinct().ToList();
        }

        public static IList<StructureDefinition> Merge(
            IEnumerable<StructureDefinition> source,
            IEnumerable<StructureDefinition> configured)
        {
            List<StructureDefinition> overrides = (configured ?? Enumerable.Empty<StructureDefinition>()).ToList();
            ThrowIfInvalid(overrides);
            List<StructureDefinition> result = (source ?? Enumerable.Empty<StructureDefinition>()).Select(Clone).ToList();
            var changed = new HashSet<StructureDefinition>();
            foreach (StructureDefinition row in overrides)
            {
                var names = new HashSet<string>(new[] { row.CanonicalName }.Concat(Names(row.Aliases))
                    .Select(StructureAliasResolver.NormalizeName));
                List<StructureDefinition> matches = result.Where(item =>
                    !string.IsNullOrWhiteSpace(row.StructureId)
                        ? string.Equals(row.StructureId, item.StructureId, StringComparison.OrdinalIgnoreCase)
                        : names.Contains(StructureAliasResolver.NormalizeName(item.CanonicalName))).ToList();
                if (matches.Count > 1 || (matches.Count == 1 && changed.Contains(matches[0])))
                {
                    throw new FormatException("Mehrdeutige Zuordnung zur Constraint-Quelle: " + row.CanonicalName + ".");
                }
                StructureDefinition replacement = Clone(row);
                if (matches.Count == 1)
                {
                    StructureDefinition original = matches[0];
                    replacement.StructureId = original.StructureId;
                    replacement.Active = original.Active;
                    replacement.Codes = original.Codes;
                    replacement.DicomTypes = original.DicomTypes;
                    // Source names can still be referenced by constraint rows without IDs.
                    replacement.Aliases = Names(replacement.Aliases).Concat(new[] { original.CanonicalName })
                        .Where(name => !string.Equals(name, replacement.CanonicalName, StringComparison.OrdinalIgnoreCase))
                        .Distinct(StringComparer.OrdinalIgnoreCase).ToList();
                    result[result.IndexOf(original)] = replacement;
                }
                else
                {
                    result.Add(replacement);
                }
                changed.Add(replacement);
            }
            ThrowIfInvalid(result);
            return result;
        }

        public static string Serialize(IEnumerable<StructureDefinition> definitions)
        {
            var rows = (definitions ?? Enumerable.Empty<StructureDefinition>()).ToList();
            ThrowIfInvalid(rows);
            var document = new AliasDocument
            {
                Schema = Schema,
                NomenclatureSource = NomenclatureSource,
                Structures = rows.Select(row => new AliasRow
                {
                    StructureId = row.StructureId,
                    CanonicalName = row.CanonicalName.Trim(),
                    Aliases = Names(row.Aliases).ToList(),
                    Laterality = row.Laterality ?? string.Empty,
                    SideAliasesLeft = Names(row.SideAliasesLeft).ToList(),
                    SideAliasesRight = Names(row.SideAliasesRight).ToList()
                }).ToList()
            };
            return JsonConvert.SerializeObject(document, Formatting.Indented);
        }

        public static IList<StructureDefinition> Deserialize(string json)
        {
            AliasDocument document;
            try
            {
                document = JsonConvert.DeserializeObject<AliasDocument>(json,
                    new JsonSerializerSettings { MissingMemberHandling = MissingMemberHandling.Error });
            }
            catch (JsonException exception)
            {
                throw new FormatException("Alias-JSON ist ungültig: " + exception.Message, exception);
            }
            if (document == null || document.Schema != Schema || document.Structures == null)
            {
                throw new FormatException("Erwartetes Alias-Schema: " + Schema + ".");
            }
            var rows = document.Structures.Select(row => row == null ? null : new StructureDefinition
            {
                StructureId = row.StructureId,
                CanonicalName = row.CanonicalName,
                Active = true,
                Aliases = Names(row.Aliases).ToList(),
                Laterality = row.Laterality,
                SideAliasesLeft = Names(row.SideAliasesLeft).ToList(),
                SideAliasesRight = Names(row.SideAliasesRight).ToList()
            }).ToList();
            ThrowIfInvalid(rows);
            return rows;
        }

        public static IList<StructureDefinition> Import(string path)
        {
            if (string.Equals(Path.GetExtension(path), ".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                return ImportCatalog(new ExcelConstraintSource().Load(path));
            }
            string json = File.ReadAllText(path, Encoding.UTF8);
            string schema;
            try { schema = (string)JObject.Parse(json)["schema"]; }
            catch (JsonException exception) { throw new FormatException("Alias-JSON ist ungültig.", exception); }
            return schema == RefDbJsonConstraintSource.SupportedSchema
                ? ImportCatalog(new RefDbJsonConstraintSource().Load(path))
                : Deserialize(json);
        }

        public static void Save(string path, IEnumerable<StructureDefinition> definitions)
        {
            if (string.IsNullOrWhiteSpace(path) || !string.Equals(Path.GetExtension(path), ".json", StringComparison.OrdinalIgnoreCase))
            {
                throw new ArgumentException("Eine separate JSON-Datei für Aliase ist erforderlich.");
            }
            string json = Serialize(definitions);
            // Never overwrite a RefDB export (or any other JSON document).
            if (File.Exists(path)) Deserialize(File.ReadAllText(path, Encoding.UTF8));
            string temporary = path + ".tmp." + Guid.NewGuid().ToString("N");
            try
            {
                File.WriteAllText(temporary, json, new UTF8Encoding(false));
                if (File.Exists(path)) File.Replace(temporary, path, path + ".bak", true);
                else File.Move(temporary, path);
            }
            finally
            {
                if (File.Exists(temporary)) File.Delete(temporary);
            }
        }

        private static IList<StructureDefinition> ImportCatalog(ConstraintCatalog catalog)
        {
            if (catalog.Issues.Any(issue => issue.IsFatal))
            {
                throw new FormatException(string.Join(Environment.NewLine, catalog.Issues.Where(issue => issue.IsFatal).Select(issue => issue.Message)));
            }
            // Preserve imported source names for review. Validation happens before export/application.
            return catalog.Structures.Select(Clone).ToList();
        }

        private static IEnumerable<string> AllNames(StructureDefinition row)
        {
            return new[] { row.CanonicalName }.Concat(Names(row.Aliases))
                .Concat(Names(row.SideAliasesLeft)).Concat(Names(row.SideAliasesRight)).Distinct();
        }

        private static IEnumerable<string> Names(IEnumerable<string> values)
        {
            return (values ?? Enumerable.Empty<string>()).Where(value => !string.IsNullOrWhiteSpace(value))
                .Select(value => value.Trim()).Distinct(StringComparer.OrdinalIgnoreCase);
        }

        private static StructureDefinition Clone(StructureDefinition row)
        {
            return new StructureDefinition
            {
                StructureId = row.StructureId, CanonicalName = row.CanonicalName, Active = row.Active,
                Laterality = row.Laterality, Aliases = Names(row.Aliases).ToList(),
                SideAliasesLeft = Names(row.SideAliasesLeft).ToList(), SideAliasesRight = Names(row.SideAliasesRight).ToList(),
                Codes = Names(row.Codes).ToList(), DicomTypes = Names(row.DicomTypes).ToList()
            };
        }

        private static void ThrowIfInvalid(IEnumerable<StructureDefinition> rows)
        {
            IList<string> errors = Validate(rows);
            if (errors.Count > 0) throw new FormatException(string.Join(Environment.NewLine, errors));
        }

        private sealed class AliasDocument
        {
            [JsonProperty("schema")] public string Schema { get; set; }
            [JsonProperty("nomenclature_source")] public string NomenclatureSource { get; set; }
            [JsonProperty("structures")] public IList<AliasRow> Structures { get; set; }
        }

        private sealed class AliasRow
        {
            [JsonProperty("structure_id", NullValueHandling = NullValueHandling.Ignore)] public string StructureId { get; set; }
            [JsonProperty("canonical_name")] public string CanonicalName { get; set; }
            [JsonProperty("aliases")] public IList<string> Aliases { get; set; }
            [JsonProperty("laterality")] public string Laterality { get; set; }
            [JsonProperty("side_aliases_left")] public IList<string> SideAliasesLeft { get; set; }
            [JsonProperty("side_aliases_right")] public IList<string> SideAliasesRight { get; set; }
        }
    }
}
