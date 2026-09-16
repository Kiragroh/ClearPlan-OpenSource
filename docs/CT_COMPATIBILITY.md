# CT compatibility approvals

[Documentation](index.md) · [Configuration guide](configuration-guide.md)

`CtCompatibilityJsonPath` in `settings.ini` selects a local configuration. The public
`CtCompatibility.json` starter has no approvals and never produces a CT compatibility pass.

Use **Settings → Configurations & versions**, selecting the CT approval document, to
import an existing configuration, edit its JSON, save a revision with a reason, or restore
an older version. Save the settings to activate the selected immutable revision. The
current review is then rebuilt without reopening the patient; the same detached findings
feed GUI and reports. A path can also be set under **Paths & sources**.

Synthetic format example (not a clinically approved device):

```json
{
  "schemaVersion": 1,
  "combinations": [
    {
      "id": "synthetic-example",
      "enabled": false,
      "manufacturer": "SYNTHETIC",
      "model": "CT Test",
      "serialNumber": "000123",
      "calibration": "HU Test"
    }
  ]
}
```

Every field shown is required. Manufacturer, model, serial number and HU-calibration
identifier must match **the same enabled entry**. Matching trims only leading/trailing
whitespace and ignores case using ordinal comparison. Internal whitespace, punctuation,
serial-number leading zeroes and complete calibration names remain significant. There
are no wildcards, manufacturer-only approvals, model families or inferred site defaults.
`enabled` must be a JSON boolean and identifiers must be JSON strings. Duplicate keys,
entry IDs or complete tuples, unknown fields, unsupported schema versions, wildcards,
missing values and files over 128 KiB are rejected. A network read times out after two
seconds; the worker reads configuration bytes only and never receives ESAPI objects.

Results:

- Exact enabled complete tuple: pass, with approval entry ID and observed tuple.
- Complete nonmatching tuple with configured active approvals: fail.
- Missing/invalid observations or absent, invalid, empty or wholly disabled approvals:
  not evaluated, never green.

## Deliberately narrow legacy boundary

The private legacy `ErrorCalculator` uses code `19` for its catch-all CT/HU warning.
Its message contains the complete observed `Series.ImagingDeviceSerialNo`,
`Series.ImagingDeviceModel`, `Series.ImagingDeviceManufacturer` and
`Series.ImagingDeviceId` (HU calibration) tuple. The adapter accepts only the raw code
`19` and its full anchored four-field message grammar. Loose strings, extra trailing
content, incomplete fields and generated duplicate IDs are not an approval input.
The original finding is retained as provenance and the upstream ErrorGrid is unchanged.

Other private legacy CT branches (including their existing hard-coded passes), all
other legacy checks, native Eclipse warnings and the public starter calculator remain
unchanged. This feature is not a replacement of the complete CT-check family. A future
migration of that family should use structured native observations for every plan,
including plan sums; do not extend the adapter using fuzzy message searches.

The existing optional PlanCheck inclusion policy remains separate: excluding a family
is still documented as exclusion, not a successful evaluation. No CT, calibration,
plan, ARIA data or clinical source report is modified. Configuration/test approval is
not clinical validation or commissioning.
