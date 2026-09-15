# Structure names and local aliases

ClearPlan can apply a separate, versioned alias JSON without changing a constraint workbook, a RefDB export, a prescription mapping, or a TPS structure. `StructureAliasesJsonPath` lives in `[Paths]` in `settings.ini`; relative paths resolve from the application build folder. The default is blank, preserving the selected source's existing mappings, including local installations.

## Configure in the clinical settings window

The **Strukturnamen und Aliase** table displays the active source mappings. Use **Quelle übernehmen** to start from the selected source without existing overrides, or **Import JSON / Excel** to read an alias document, a RefDB constraint export, or an existing ClearPlan constraint workbook (`Structures` worksheet in the established catalog schema).

Edit the canonical name, explicit aliases separated by `|`, laterality and side aliases. The optional source ID is retained read-only when imported. **Aliase prüfen** checks duplicate identities, normalized name collisions, unknown laterality and cross-side aliases, including conflicts with the selected source. It does not certify TG-263 nomenclature or anatomical equivalence; a qualified local review remains necessary.

**Als versionierten Entwurf übernehmen** stages the aliases under **Konfigurationen & Versionen**. Enter a change reason and choose **Prüfen & als Version speichern**; **Einstellungen speichern** then activates the committed path and reloads the catalog. **JSON exportieren …** exports a copy without changing active settings. No import writes back to its source. Exporting over an existing valid alias document creates a `.bak`; other JSON schemas (including RefDB exports) cannot be overwritten by the alias writer.

## Application and safety

- A row matches an existing source definition by explicit source ID, or, without an ID, by the source canonical name appearing in the configured canonical name/aliases. Multiple matching rows are rejected.
- Matched definitions retain their source ID, active state, code metadata and DICOM types. Configured name and side lists replace the source's name lists. The previous canonical name remains as an alias so source constraint rows without IDs still resolve. Unlisted source definitions remain unchanged.
- Alias documents contain no dose thresholds. Clinical table selection, prescription labels, units, numerical goals and variations are not modified.
- Missing, malformed or conflicting explicitly configured alias files make the catalog unusable. ClearPlan does not silently switch the clinical constraint source to hide an alias error. Network alias reads use the existing bounded source timeout and cannot later change a live catalog after a timeout.
- Matching ignores case, spaces and underscores. TG-263 semantic markers (`~`, `-`, `+`, `^`, `!`, `=`, `/`) are retained. A partial organ is not a full organ; a cropped target is not a numbered target. Aliases are exact normalized names, never regex or fuzzy rules.
- Paired individual organs (`L/R`) and a combined organ (`R+L`) are different concepts. Do not map both sides to one ambiguous general alias, or map a canal to the spinal cord.

## Public example and provenance

`ConstraintTemplates/TG263_StructureAliases.example.json` is an opt-in example, not an automatically installed clinical policy. Its alternative names come only from the repository's existing public Starter CSV fixtures. The example separates spinal cord from spinal canal, and individual lungs from combined lungs. It intentionally does not guess local target or parotid conventions.

Canonical names were checked against the [AAPM TG-263 resources](https://www.aapm.org/pubs/reports/RPT_263_Supplemental/default.asp) and their [Eclipse structure templates](https://www.aapm.org/pubs/reports/RPT_263_Supplemental/EclipseStructureTemplates.zip). `Lungs` and the semantic/side conventions are stated in [TG-263, section 7.2](https://www.aapm.org/pubs/reports/RPT_263.pdf). Verified 7 September 2026. This does not imply AAPM endorsement.

The existing public Starter constraint workbook remains an illustrative software test/reference catalog. Its values have not been replaced or relabeled as Timmermann-derived guidance. No claim of clinical validation, institutional endorsement, or verified Timmermann provenance is made. Local UKL prescription/RefDB tables are not distributed in this repository.

The JSON schema is `ClearPlan.structure_aliases.v1`. Each `structures` entry supports `canonical_name`, optional `structure_id`, `aliases`, `laterality`, `side_aliases_left`, and `side_aliases_right`. Omit `structure_id` in portable mappings unless matching one specific source catalog is intentional.
