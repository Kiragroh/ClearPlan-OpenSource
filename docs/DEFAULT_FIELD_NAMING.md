# Default field-name nomenclature

The native settings page includes **Konfigurationen & Versionen → Default · Feldnamen**.
This controls read-only suggestions in ClearPlan. It does not rename fields in the TPS
or alter the separately deployed PlanFieldNamer script.

1. Select the configured JSON source and enter a change reason. Import explicitly;
   this creates a managed copy without modifying the original.
2. Edit the JSON and choose **Prüfen & als Version speichern**.
3. Save the settings to activate that exact revision, then reload the plan review.
   Merely creating a revision does not change the active configuration.
4. To return to an older configuration, select it in the version history and restore
   it with a reason. This creates a new revision; save settings to activate it.

The initial file is `FieldNamingRules.json`; its path is
`[Paths] FieldNamingRulesJsonPath` in `settings.ini`. Managed versions use the
configured `ConfigurationDirectory`.

| JSON field | Meaning | Shipped value |
| --- | --- | --- |
| `enabled` | Enable suggestions; false leaves them not evaluated | `true` |
| `arcOrder` | Each arc component exactly once | `["angles", "table", "direction"]` |
| `staticOrder` | Each static component exactly once | `["angles", "table"]` |
| `partSeparator` | Separator between components | one space |
| `angleSeparator` | Start/stop angle separator | `-` |
| `tablePrefix` | Nonzero couch-angle prefix | `T` |
| `clockwiseToken` | Clockwise direction | `UZ` |
| `counterClockwiseToken` | Counterclockwise direction | `GUZ` |
| `staticDuplicateToken` | Token before duplicate suffix for static fields | `UZ` |

The schema version is `1`. An arc from 120° to 30°, couch 300°,
counterclockwise becomes **120-30 T300 GUZ**. The couch component is penultimate;
direction is last. Repeated identical arcs become **120-30 T300 GUZa** and
**120-30 T300 GUZb**. A zero couch angle is omitted as in the existing convention.
Duplicate suffixes remain attached to the direction token if the component order
is customized. The existing field-ID prefix and treatment-order sequence policy is
unchanged; ID and name conformity remain separate results.

Missing, disabled or invalid files do not silently regenerate defaults or claim
conformity. The UI shows not evaluated and leaves suggestions empty. Supported
tokens, bounded file size and component ordering are validated before saving.
These checks establish file validity, not clinical appropriateness. Test a site's
chosen convention against representative plans before operational use.

Version history records the supplied Windows account, UTC time, change reason and
SHA-256. It is local provenance, not an authenticated approval signature. See
[configuration history](CONFIGURATION_HISTORY_STORE.md) for storage and recovery limits.
