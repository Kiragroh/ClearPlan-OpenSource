# Managed configuration history

`ClearPlan.Core.Configuration.ConfigurationHistoryStore` is an ESAPI-free,
file-based store for configuration documents. The application supplies an explicit
managed root, document-specific validation, the Windows actor and a change reason.
The store has no patient API, external source path API, database dependency or shell
integration. Use it only for configuration; do not put patient data, secrets or
clinical free text into files, reasons or metadata.

## File contract

For a key such as `aliases`, a JSON document uses this layout:

```text
<explicit managed root>/aliases/
    .lock
    current.json
    head.json
    revisions/
        00000001/content.json
        00000001/metadata.json
        00000002/content.json
        00000002/metadata.json
```

`current.json` is the internal managed working file. `head.json` lists committed revisions.
`GetManagedPath` returns the verified **immutable committed snapshot** path, not
`current.json`. Active settings therefore pin a committed revision until the user
explicitly applies a newly saved/restored revision. Another save or an interrupted
working-file update cannot change the bytes on an already activated path.
Revision snapshots are never overwritten by the store. Each revision records key,
category, extension, revision number, SHA-256, UTC time, actor, reason, action, and
the source revision for a restore. History is returned newest first.

The first import has action `original`, subsequent imports `import`, saves `edit`,
and restores `restore`. A restore appends a new revision containing the selected
historical bytes; it never moves the history backwards or deletes later versions.

Use distinct keys for distinct formats (for example, `constraints-excel` and
`constraints-refdb`). An existing key cannot change category or extension.
Keys are lower-case ASCII letters/digits/underscores/hyphens, 1–64 characters,
excluding Windows device names. Extensions are limited to `.json`, `.ini`, `.xlsx`.
Documents must contain at least one byte and at most 32 MiB (33,554,432 bytes).
The history manifest is bounded at 16 MiB; it is never silently truncated or pruned.
Managed paths containing existing reparse points/junctions/symlinks are rejected.

## Caller sequence

1. Read only the external source explicitly chosen by the user. Bound that read
   before allocating memory, and apply its document-specific validation.
2. Call `Import(key, category, extension, bytes, actor, reason, expectedRevision,
   validate)`. Use expected revision `0` only for a new key. External sources are
   never a save target: the store only accepts their copied bytes.
3. After success, retrieve `GetManagedPath(key)` and use that immutable committed
   snapshot path in the candidate settings. Do not persist or activate new settings
   before the import succeeds; applying Settings explicitly activates the revision.
4. Save editor contents with `Save(..., expectedRevision, validate)`. Supply the
   revision loaded with the editor, not a freshly fetched version that would hide
   a concurrent edit.
5. Call `Restore(..., sourceRevision, ..., expectedRevision, validate)` only after
   selecting a historical version and recording a reason.

The validation callback throws on invalid content and must validate without
mutating its input. It runs before the commit; it receives a defensive copy.
Restore always checks the stored SHA-256 and snapshot metadata before validation.
UI code should derive `actor` from the current Windows identity, not an editable
name field. Actor metadata is local provenance, **not authenticated attribution**.

External editors, including Excel, should edit a separate working copy. Submit its
validated bytes explicitly through `Save`/`Import`; do not launch an editor against
`current.*` or a revision snapshot. An out-of-band change to `current.*` is detected
and further saves/activation fail closed instead of overwriting it.
`CreateTemporaryFile` creates bounded GUID-named files only in `.editing` or
`.validation` below the explicit managed root, with the same reparse-point guards.
`DeleteTemporaryFile` removes only those generated paths, never arbitrary source
files or whole directories. Read/list operations never create a missing store.

## Atomicity, failures and maintenance

Per-key `FileShare.None` locks reject another concurrent reader/writer rather than
waiting indefinitely. Expected-revision comparison rejects stale editor contents.
The current file's SHA-256 must match committed history before a change is allowed.
An unavailable share, denied access or unsupported file operation is an error, not
a reason to fall back to an untracked save path.

The commit stages and flushes all files, publishes an immutable snapshot directory,
atomically replaces `current.*`, then atomically replaces `head.json` in the same
directory. A caught failure before head publication restores the previous working
file and does not report success. If this rollback also fails, the exception names
that condition and the rollback file is retained for administrative inspection.

The two replaced files are **not a filesystem-wide transaction**. Sudden process
termination, power loss, or ambiguous network failures between replacements can
leave a working-file/head mismatch or an unpublished snapshot. Subsequent reads
and writes detect a current/head mismatch and stop; they do not guess which file
to trust. Unpublished snapshot directories are not listed/restored as committed
history and are never overwritten (revision numbering can therefore contain gaps).

On an integrity error, stop editing, preserve the complete key directory, and have
an administrator compare the working file, head, snapshot checksums and any rollback
file against an independent backup. Do not delete `head.json` or manually change a
checksum to make the warning disappear. Normal user rollback is the append-only
`Restore` operation, not manual filesystem edits. Back up the complete managed root
using the site's normal access controls and backup policy; do not back up only the
current file. No retention deletion is performed by this implementation.

These hashes detect accidental content/metadata changes relative to the manifest;
someone with write access can rewrite both data and manifest. This is not a
tamper-proof audit trail, regulatory compliance feature, or clinical validation.
Network locking/replace/durability semantics must be verified on the intended share.

## Focused synthetic verification

`ClearPlan.Core.Tests.exe ConfigurationHistory` covers external-source preservation,
Unicode provenance, append-only restore, stale-write/current-hash conflicts, unsafe
names and bounds, callback rejection, corrupted snapshot content/metadata, malformed
or missing manifests, locked-head rollback and exclusive writer locking. Tests use
new temporary directories and synthetic contents only; no clinical runtime is used.
