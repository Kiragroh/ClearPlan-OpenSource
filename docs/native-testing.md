# Native testing and the Runner

[Documentation](index.md) · [Developer handoff](developer-handoff.md) · [ESAPI Runner Hub](https://github.com/Kiragroh/ESAPI-Runner-Hub)

ClearPlan's calculation and reporting code does not depend on ESAPI Runner Hub.
The Runner provides the configured launch route used for the local native
integration tests. Another suitable authorized ESAPI host can serve that role.

## Which tests need it?

| Work | Runner needed? | What the result establishes |
|---|---|---|
| Portable ClearPlan C# tests | No | Core behavior with synthetic inputs and source-boundary contracts |
| Standalone simulator and patient-free settings capture | No | Shared presentation/configuration behavior without a patient session |
| Native ESAPI acquisition, context switching, CT/BEV, and report checks | An authorized native host is needed; Runner is our route | Behavior in the selected installed ESAPI environment |
| Repeat an exact context through the documented Citrix request workflow | Yes, for this workflow | The configured script was dispatched on the identified server for the requested context |
| Clinical acceptance | No particular launcher suffices | Requires independently reviewed rules, measurements, workflows, and local commissioning |

The 495 portable tests at ClearPlan code commit
`98aed7bc2fb3ecbced7cc2a520f691da36c53bfa` are not Runner tests or 495 clinical plans.
Runner regressions have separate evidence and counts.

## Why include the launch route in a technical note?

A portable calculation test cannot prove that a native plan opens, the expected
course/plan is resolved, the installed API supplies the required data, or a GUI
and report render in the actual application-server environment. The context and
loaded build therefore belong to the test setup.

A shared published Citrix entry point can launch independently versioned tools
without adding a publication for each script. Explicit user-scoped requests
avoid reliance on client-side command-line forwarding and on whichever patient
was most recently selected. Separate child processes isolate target failures
from the catalogue. They do not guarantee clinical safety or arbitrary parallel
ESAPI access.

ClearPlan can also run through the normal Eclipse script entry point or another
compatible locally authorized host. The Runner is an operational convenience and
a documented reproduction route, not a new requirement for the core library or
for other TPS adapters. RayStation needs its own native acquisition and test setup.

## Reproduce a native check

1. Use an authorized ESAPI-capable workstation or published application server.
   Opening an executable from a network share does not relocate its process to
   that server. Supply licensed dependencies locally; do not distribute them.
2. Configure a specific ClearPlan entry, target path and read-only script host.
   Record the Runner, host and ClearPlan versions/hashes separately, along with
   configuration revisions. An application label alone is insufficient.
3. Select an explicit patient/course/plan, or create a protected user-scoped
   request using the Runner's documented interface. Never put real identifiers
   in a public example, issue, command transcript, or repository.
4. Execute sequentially for a controlled test series. Preserve the matching
   request/result locally and verify context acquisition, application-specific
   probe outcomes, report contents and image availability separately.
5. Inspect the rendered output. A successful host exit is not proof that all
   optional images were available or that a clinical finding was correct.

See the Runner's [native-test rationale](https://github.com/Kiragroh/ESAPI-Runner-Hub/blob/main/docs/NATIVE_TESTING.md)
and [request contract](https://github.com/Kiragroh/ESAPI-Runner-Hub/blob/main/docs/CONTEXT_DEBUGGING.md).
Clinical capture helpers and patient-bound sidecars are local test tooling, not
public datasets. Access control and retention of readable request files are the
deployment's responsibility; a SID check does not encrypt them.

## Boundaries

Keep ClearPlan on the read-only host. The Runner also supports separately governed
write-enabled applications; that capability is not used to make these reviews or
tests writable. No launcher replaces script approval, Windows permissions, TLS
validation, endpoint protection, or institutional commissioning. Optional ARIA
upload is a separate external write and is excluded from native no-send tests.

Publish a source reference and patient-free reproduction instructions. Keep
patient requests, private rule matrices, raw captures and clinical reports on
protected storage. A masked screenshot is not automatically publication-cleared.
