# Optional ARIA report upload

[Documentation](index.md) · [Configuration guide](configuration-guide.md)

The native workspace has an **An ARIA senden** action next to PDF and HTML.
It creates the current single-plan report, including the current visibility and
optional field-start BEV settings. It does not upload an arbitrary older PDF.

ESAPI plan access stays read-only. Creating a DocumentReference is a separate,
explicitly confirmed write to the ARIA patient record. It does not approve the
plan or change dose, structures, treatment fields, or clinical-goal settings.

## First-time setup: choose and commission a document route

FHIR/VAIS is not installed or enabled by ClearPlan. If the clinic does not yet
have it, the TPS/ARIA administrator should open a **Varian support ticket via
[MyVarian](https://www.myvarian.com/login)** requesting installation/configuration
of the ARIA API/FHIR service for PDF DocumentReference create and readback. Ask
Varian to confirm compatibility with the installed ARIA version, local entitlement,
server prerequisites, the current ARIA API Reference Guide/Implementation Guide,
VAIS client registration, and the required permissions. MyVarian product documents
may require an account and the appropriate product entitlement; this is not an
automatic ClearPlan installation step.

Create a **local document-upload profile** with your ARIA administrator. Choose
one route explicitly:

| Route | What the clinic configures |
|---|---|
| ClearPlan's **An ARIA senden** action | The FHIR connection JSON below, personal-user Practitioner mapping, document type/category and a separately authorized test upload/readback |
| Existing **Webservice / direct upload** workflow | A locally commissioned ARIA Oncology Services/Gateway document profile, its authentication, patient association, TemplateName and personal AuthoredBy mapping; use the exported PDF with that workflow |
| **eDocPrinter** | A local printer/import profile for the ClearPlan report, including unambiguous patient-ID extraction, document type, template/author mapping and destination; verify the resulting patient document after an authorized test |

Webservice/Gateway and eDocPrinter are alternatives managed by the clinic, not
fallback transports implemented by ClearPlan's FHIR button. Do not put Gateway
endpoints into the FHIR JSON. An exported PDF alone does not establish correct
patient, template or author routing. Keep upload disabled until the selected
profile has been commissioned; do not use direct SQL for document writes.

## Site configuration

`Paths.AriaUploadConfigJsonPath` in `settings.ini` points to external JSON. An empty
path disables the action. The path is editable in Settings; the JSON can be
imported, validated, versioned and restored in the configuration workspace.
Start with `AriaUpload.example.json`, which is disabled and contains no site data.

Request and configure `system/DocumentReference.cruds`, `system/Patient.rs`,
`system/Organization.rs`, `system/ValueSet.rs` and **`system/Practitioner.rs`**.
Existing local JSON profiles must add the Practitioner scope and the VAIS client
must actually be granted it; editing a scope string alone grants no permission.
Never share a client secret in a support-ticket body or commit it to the project.

Configure HTTPS FHIR and token endpoints, OAuth client ID and scope, the exact
provider reference and document type. If no provider is configured, exactly one
active provider must resolve. Type resolution uses the provider-specific ARIA
ValueSet and exact code/display matching, not fuzzy selection. The site's existing
PDF import configuration is a useful starting point; verify it locally.
`CategoryCode` must match the category ARIA actually stores (`PDF` or `TIF`, as
configured by the site). ClearPlan does not silently accept a different category;
the attachment must remain `application/pdf` in either case.

`DocumentDateSafetyMinutes` defaults to **0** and accepts whole minutes from 0 to
60. It is an explicit site compatibility setting for servers that reject a
document date with `FUTURE_DATE_TIME`, not a timezone conversion. The ARIA
document date is `min(report creation, current UTC time - configured margin)`.
The report's captured creation instant, as printed in the PDF header, remains
the attachment `creation` value. PDF bytes and writer metadata are not changed.
If the two dates differ, the confirmation shows both UTC values and the configured
margin. Verify workstation/server clocks and site requirements before changing
this setting. If the server also rejects attachment creation, investigate that
separately; ClearPlan does not silently backdate the attachment or retry the POST.

Keep the client secret in the configured environment variable or external dotenv
file. The JSON accepts a credential file reference, never a `ClientSecret` value.
An absolute credential path avoids changing its meaning when configuration is
imported into a versioned folder. Do not put credentials in Git or shared examples.

Normal OS certificate validation is the default. For explicitly approved private
installations, separate SHA-256 certificate bindings can be supplied for token and
FHIR endpoints. Verify the fingerprints through your IT team before deployment.
Binding does not accept changed/expired certificates, hostname mismatch, or
revocation errors. It permits only the exact valid leaf certificate with an
otherwise untrusted private root. There is no automatic pin renewal or blanket
TLS-disable switch. Connections use TLS 1.2 and do not follow redirects.

## TemplateName and personal AuthoredBy

The independent ARIA TemplateName is **`PQM_<active plan name>`**, using the exact
Eclipse plan ID shown as the active plan name (without the course label or PDF
filename). It is sent in
`http://varian.com/fhir/v1/StructureDefinition/documentreference-templateName`
as `valueString`; it is not inferred from the attachment title. Unicode and
punctuation are retained. Missing, unsafe or overlong names block preparation
instead of being silently truncated or replaced.

`DocumentReference.author[0].reference` is a resolved **Practitioner** for
`ScriptContext.CurrentUser.Id`, captured on the ESAPI owner thread. For a
domain-qualified identity, lookup and exact case-insensitive verification use
`http://varian.com/fhir/identifier/Practitioner/UserName`; an unqualified ARIA
identity uses the exact, case-sensitive `.../Practitioner/Id` system. The installed
IG's Practitioner search and example resource document these identifiers. No
Windows domain is invented or stripped. `CurrentUser.Name` is display text, not
an identity ([official ESAPI User reference](https://docs.developer.varian.com/api/17.0/VMS.TPS.Common.Model.API.User.html)).

Zero, inactive, conflicting or multiple matching Practitioners block sending
with an actionable message. Ask the ARIA administrator to check the personal
user-to-staff/Practitioner association and read/search permission. There is no
fuzzy name search, hard-coded author, or fallback to the OAuth service account.
The service account still authenticates the HTTP request and may appear in ARIA
technical/audit fields; this is distinct from the document's **AuthoredBy**.
No approval/signature/authenticator is added.

The confirmation lists TemplateName and the interactive author/reference. A plan,
patient or user-context change invalidates preparation. Legacy request constructor
signatures remain available for inspection, but the client refuses to send a
request without explicit plan/author binding, before OAuth or POST.

## Send and verify

The action is unavailable in synthetic/sandbox or privacy mode. It resolves the
active patient by exact identifier, prepares immutable PDF bytes, then displays
patient, plan, TemplateName, personal author, provider, document type and SHA-256 for confirmation. Changing
context while preparing invalidates the request. A second click cannot start a
parallel upload.

The upload uses the local ARIA DocumentReference profile, `status=current`,
`docStatus=preliminary`, PDF attachment and `documentLocation=file-server`. It
never adds an approval signature. Server responses are not treated as approval.
Rejected uploads expose the numeric HTTP status and, when present, the allowlisted
business-rule code `FUTURE_DATE_TIME` with a fixed actionable message. Raw server
diagnostics, unknown codes and clinical response text are not displayed or logged.

A successful POST is followed by a separate readback of patient, provider,
document type, category, current/preliminary status and PDF content type. Exactly
one matching TemplateName extension and exactly one matching Practitioner author
are required: missing, different, duplicated or additional authors/templates do
not count as verified, even when the PDF bytes match. The
document date must match within 3 ms, allowing database timestamp rounding.
Attachment creation, if returned, must match the captured report creation within
the same tolerance. If ARIA omits it, the result and GUI explicitly say that it
was not checked; this is not evidence that the server preserved that metadata.

A returned title must match exactly. If the title is omitted, the nonempty report
description must match exactly, and full verification still requires matching
PDF bytes or a returned fingerprint. Metadata-only verification never counts as
a verified upload. When returned, size, inline bytes and hash are all checked;
contradictory values fail verification.

`AttachmentReadbackRoots` defaults to an empty array, disabling filesystem
readback. Sites where ARIA returns only a stored PDF UNC path can explicitly
configure up to eight approved UNC roots. Only descendants of those roots are
read, using case-insensitive path boundaries. HTTP URLs, `file:` URLs, local or
device paths, traversal, alternate data streams and ambiguous Windows names are
rejected. Directory and file handles are checked for Windows-reported reparse
points and held against rename while reading; the PDF is opened read-only and
against concurrent writes. Reads are bounded to 10 seconds and 20 MiB. Returned
bytes must have exactly the prepared PDF's length and SHA-256. No attachment URL
is followed over HTTP and no file is written. Configure least-privilege read
access to the intended document repository; do not approve a broad unrelated
share. Public examples contain no institution-specific roots.

## Uncertain results and storage

PDFs and append-only receipts are kept in `ReportsDirectory/AriaUploads`. Use an
appropriately protected location. Receipt filenames are hashed target/plan
bindings, not patient identifiers. Receipt contents and PDFs are still confidential.

The intent is persisted before the single POST. Pending, unverified Sent, or
Uncertain results block another upload for that target/plan, including after a
Runner restart. There is no POST retry after timeout. Reconcile the receipt with
ARIA before an administrator permits another attempt; do not delete receipts to
retry blindly. A verified previous upload is disclosed before a consciously new
report can be confirmed.

If an accepted create response has a validated same-origin document Location but
an unreadable response body, that reference is retained with the uncertain
receipt. The GUI can verify this existing resource with GET only; it does not
resend the document. Location query strings are never followed or forwarded:
only the validated canonical `DocumentReference/id` is used. Conflicting body and
Location identifiers remain unconfirmed rather than selecting one arbitrarily.
An unexpected response target still makes the POST result uncertain. A separately
validated Location may remain as a GET-only recovery candidate after the bounded
body consistency check; a body-only reference cannot substitute for it in this
failure path. Conflicting identifiers discard the candidate. Only the independent
identity and PDF readback can establish verification; this does not establish the
cause of an unexpected transport response target or authorize a repeated POST.
Uncertain create results expose a fixed processing-stage code (for example,
`response-target`, `create-location` or `create-body`) for troubleshooting. These
codes contain no URLs, identifiers, exception text or server response body.

Filesystem work is awaited for at most ten seconds per operation. A blocked SMB
call may finish later, but cannot start an upload or retain the UI. Late file
leases are disposed in the background. The report renderer runs on the native STA;
no ESAPI object is passed to the background filesystem or HTTP workers.

## Verification scope

Synthetic unit tests cover payload identity, exact interactive-author lookup,
TemplateName and author readback, legacy unbound-write rejection, pagination boundaries,
single-write behavior, certificates, credentials, sparse readback metadata,
timestamp integrity, constrained attachment paths and durable receipts. File
reader tests use a disposable local synthetic PDF, never an institutional share.
The native GUI probe checks visibility, privacy and sandbox guards without
performing a document upload. Local interoperability and actual write/readback
evidence must be recorded separately; unit tests are not clinical commissioning.
