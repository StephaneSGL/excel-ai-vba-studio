# ADR: safe XLSM and VBA write-back

Status: Accepted
Date: 2026-07-27

## Context

An `.xlsm` file combines normal OOXML parts with opaque content such as
`vbaProject.bin`, UserForm resources, controls, drawings, and ActiveX data.
Rebuilding the complete package with a generic spreadsheet serializer can
silently remove or alter those parts.

The extension therefore separates cell editing from VBA source editing. AI
tool output is untrusted input and goes through the same validation as a
manual edit.

## Decision

### Targeted cell and formatting edits

For a local, writable `.xlsm`, the editor computes a bounded diff containing
cell values, formulas, number formats, fonts, alignment, fills, borders,
explicit row heights and column widths, supported new conditional-formatting
rules, and explicit worksheet-wide rule clearing.

The editor records the workbook SHA-256 when the grid is loaded. A PowerShell
helper then:

1. validates protocol v2, the operation bounds, the local path, file type,
   permissions, and absence of reparse points;
2. opens the source with a writer-excluding file handle and rejects a stale
   grid baseline;
3. edits only a sibling work copy through a new invisible Excel process;
4. disables alerts, events, link updates, automatic calculation, and macros
   before opening that copy;
5. verifies the OOXML entry set, opaque control resources, and a canonical
   fingerprint of every logical VBA compound-file stream;
6. atomically replaces the source and creates a verified persistent backup in
   `.excel-ai-vba-backups`;
7. restores the original automatically from that backup if a post-replacement
   validation fails.

The timeout path targets only the process IDs owned by the transaction. It
does not kill unrelated Excel processes.

The spreadsheet is an editable VS Code custom document. Grid changes mark the
document dirty, close/revert use the standard VS Code lifecycle, and hot-exit
stores a bounded, source-hash-bound sheet-state backup in VS Code extension
storage. Recovery refuses a backup when the disk source has changed. Save and
Save As enter through the provider only while the matching custom-editor tab is
active, and a document permits one editable webview so independent split views
cannot silently overwrite each other. A disk change never reloads over unsaved
edits without an explicit user choice.

### VBA source edits

The bundled Windows helper uses pinned `pyOpenVBA` code to edit
`vbaProject.bin` directly. It does not start Excel or the VBE, execute a macro,
use AccessVBOM, or change Office/registry settings.

It can:

- update existing standard, class, document, and UserForm code-behind modules;
- add standard and class modules.

It refuses:

- new document modules or UserForm designers;
- UserForm layout or FRX changes;
- digitally signed or password-protected VBA projects;
- stale, read-only, network/device, reparse-point, invalid, or ambiguous input.

The helper writes and validates a sibling copy, verifies the persistent
backup, checks UserForm designer streams and all non-VBA OOXML payloads, then
uses an atomic replacement.

No macro-enabled write path falls back to generic OOXML serialization.

## Threat model

| Threat | Control | Residual handling |
| --- | --- | --- |
| Stale VS Code grid overwrites a newer workbook | SHA-256 captured at load, checked under a writer-excluding handle | Refuse and require reload |
| Concurrent local writer replaces the path during commit | Atomic replacement captures the exact displaced file; any unexpected version is restored without blind overwrite | Refuse the edit and retain recovery evidence |
| UNC, mapped network, device, symlink, junction, or reparse redirection | Canonical local-drive and full path-chain checks | Refuse |
| Malformed or oversized AI/tool operations | Protocol schema, UUID, size, count, worksheet/range bounds, and per-operation allowlists | Refuse |
| Workbook macros, events, prompts, or external-link updates | `AutomationSecurity = 3`, events/alerts/link updates disabled before open | Work copy only; refuse on automation failure |
| Excel changes VBA, UserForms, references, controls, or opaque resources | All logical CFB streams fingerprinted; protected OOXML resources hashed; entry set fixed | Refuse before replacement |
| Signed VBA project would lose trust | Signature detection in direct writer | Refuse VBA modification |
| Helper returns a forged path or hash | TypeScript canonicalizes the backup path and verifies backup/result hashes | Refuse |
| Excel/COM hangs | Bounded timeout and owned-PID cleanup | Report failure; source remains or is recoverable |
| Failure after atomic replacement | Verified displaced backup is restored with another capture-preserving atomic swap | Report rollback result and backup path |
| Close, revert, crash, external reload, or a stale split loses grid edits | Single editable view, provider-routed save/revert, bounded source-hash-bound hot-exit backup, and explicit reload prompt | Restore only against the exact source; otherwise retain the backup and refuse |

Assets protected include workbook data, formulas and formats, VBA sources,
references, signatures, UserForm designer/FRX streams, controls, ActiveX data,
opaque OOXML parts, local path containment, and unrelated Excel processes.

## Alternatives

- Automatic AccessVBOM or registry mutation: rejected.
- Always using VBIDE automation: rejected; policy-dependent and UI-driven.
- Generic OOXML rewrite for macro files: rejected.
- A second production C# Office Workbench host: not retained; the bounded
  PowerShell/COM path covers the current cell-edit scope with one runtime.
- Signed templates or a trusted add-in: possible future option, not
  implemented.

## Recovery

A pre-replacement failure leaves the source unchanged. A successful save
keeps a full-workbook backup beside the source. Post-replacement validation
attempts an automatic atomic restore while retaining the backup. Outside that
failure path, restoration is an explicit user action; there is no general
rollback command.

When no safe writer supports an operation, the extension refuses it and
directs the user to native Excel or the VBE.

## Verification

- `npm run validate` covers static gates, regression tests, type checking,
  production build, and package validation.
- Windows CI tests the tracked VBA helper, rebuilds it reproducibly from the
  pinned Python toolchain, requires a byte-for-byte binary match, and tests the
  rebuilt helper again.
- Manual Excel Desktop evidence on 2026-07-27: 31/31 targeted-edit checks,
  4/4 real UserForm/VBA preservation checks, 22 logical VBA streams unchanged,
  and no orphan Excel process.
- Hosted CI does not contain Microsoft Excel. The COM suite remains a Windows
  release gate on a machine with Excel installed.

## Current limits

- The source-only VBA writer edits existing UserForm code while preserving its
  designer. The separate bounded VBA Designer can create UserForms, supported
  controls, worksheet Form Control buttons, and permitted worksheet ActiveX
  controls in an existing `.xlsm`; it does not provide drag-and-drop editing.
- Worksheet ActiveX creation and `Click` binding require Excel to permit
  insertion. Third-party ProgIDs are denied unless the exact value is already
  configured in `excelAiVbaStudio.allowedCustomActiveXProgIds`.
- Explicit row heights and column widths are supported; implicit dimension
  resets remain refused. A resize that would move a protected control/drawing
  and rewrite its VML anchor also remains refused with explicit guidance.
- The five conditional-formatting presets created by the ribbon can be
  appended, and all rules on one worksheet can be cleared explicitly. Existing
  rule edits, reordering, partial deletion, expressions, and arbitrary formulas
  remain refused.
- Worksheet creation/removal, merges, and other sheet features remain refused.
- `.xls` remains read-only.

## Tracking

- [#7 — bounded VBA/UserForm transaction tool](https://github.com/StephaneSGL/excel-ai-vba-studio/issues/7)
- [#8 — safe VBA/UserForm write-back](https://github.com/StephaneSGL/excel-ai-vba-studio/issues/8)
- [#9 — targeted native XLSM editing](https://github.com/StephaneSGL/excel-ai-vba-studio/issues/9)
- [#10 — XLSM formatting parity](https://github.com/StephaneSGL/excel-ai-vba-studio/issues/10)
- [#11 — prototype consolidation](https://github.com/StephaneSGL/excel-ai-vba-studio/issues/11)
- [#20 — worksheet controls, macro assignment, and Developer window](https://github.com/StephaneSGL/excel-ai-vba-studio/issues/20)
