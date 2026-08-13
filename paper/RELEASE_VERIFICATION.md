# ClearPlan v3.1.0 release verification

Verification date: July 30, 2026

- Public release:
  https://github.com/Kiragroh/ClearPlan-OpenSource/releases/tag/v3.1.0
- Annotated tag target:
  `8bf373ebbffb17bce86ecdac29ef655b84f13e4d`
- Simulator asset:
  `ClearPlan-Simulator-v3.1.0-win-x64.zip`
- Asset size: 3,312,449 bytes
- SHA-256:
  `406424a5edf443f63d152c8a0f343eccea760693f4ddec3c06879c9beef7f55a`

## Independent checks

1. The public tag was cloned without a GitHub credential helper.
2. The clean clone built with zero compiler warnings and zero errors.
3. All 125 C# tests and 20 Python tests passed.
4. The Python-produced RayStation example snapshot passed the C# contract
   validator with zero issues.
5. All seven synthetic scenarios, responsive layout probes, five deterministic
   captures, DVH PNG, and synthetic PDF report passed the simulator smoke test.
6. Repeated package creation in the clean clone produced byte-identical
   archives.
7. Packaging with Windows PowerShell 5.1 and PowerShell 7.4.14 produced the
   same archive.
8. An unauthenticated download of both GitHub assets matched the published
   checksum, the GitHub asset digest, and the clean-clone archive.

The release verification concerns software and synthetic artifacts only. It
does not constitute clinical commissioning or a clinical performance claim.
