# Control Design — SOX Access Review Controls

This document covers the architectural rationale, design decisions, deployment requirements, and testing approach behind the access review controls script. It is written for a technical reader who wants to understand the decisions behind the implementation, not just what it does.

For a higher-level overview, see the [README](../README.md).

---

## Architecture

The script implements two independent control layers inside a single Google Apps Script file, running against a Google Sheets access review.

```
onEdit(e)
│
├── onEditAuditHistory(e)        ← Layer 2: audit layer runs first
│     ├── Read editor identity from e.user
│     ├── Read assigned reviewer from OWNER_EMAIL_COLUMN
│     ├── Derive Match Status and Edit Type
│     ├── Append full record to Audit Log tab
│     └── Write validation status to MATCH_STATUS_COLUMN in main sheet
│
└── Enforcement logic             ← Layer 1: enforcement runs after
      ├── Check if edited column is in COLUMNS_TO_WATCH
      ├── Compare editor email vs. assigned reviewer email
      ├── If unauthorized → revert cell, notify editor via toast
      └── If authorized → log reviewer timestamp
```

The audit layer is called first, deliberately. This ensures the unauthorized value and the identity mismatch are written to the audit log before the enforcement layer wipes the cell. An audit log that only captures successful edits is not a complete compliance record.

The two layers share `SHEET_NAME` and `OWNER_EMAIL_COLUMN` from the Layer 1 configuration block, but otherwise operate with no shared state and no dependency on each other. They can be modified, extended, or replaced independently.

---

## Design Decisions

### Why Google Apps Script over a standalone system

The access review process runs in Google Sheets — owned by the compliance team, shared with reviewers, reviewed by auditors. A standalone system would introduce adoption friction and audit complexity: a new tool requires explanation to external auditors and a separate evidence trail. Building the controls directly into the sheet means audit evidence is generated in the same environment auditors already work with, with no additional handoff.

### Why installable triggers, not simple triggers

Google Apps Script simple triggers do not reliably expose `e.user` — the authenticated editor's email — for all users in a Google Workspace environment. Without `e.user`, both enforcement and attribution fail silently: the script runs, finds no email, and exits without reverting or logging anything.

Installable triggers run under a specific authorized account and capture `e.user` correctly across all editors. This constraint is non-obvious and caused silent failures during early testing. It is documented explicitly in the deployment checklist and flagged in the script header to prevent recurrence on future deployments.

### Why log unauthorized attempts before the revert fires

`onEditAuditHistory()` is called as the first line of `onEdit()`, before any revert logic executes. This ordering ensures the audit log captures:

- The value the unauthorized editor attempted to enter
- The identity mismatch between editor and assigned reviewer
- The edit type as `Unauthorized Attempt`

If the audit layer ran after the revert, the cell value would already be wiped and the attempt would be partially unrecoverable from the log. The compliance record must be complete regardless of whether the enforcement succeeds.

### Why `No Match` persists in the main sheet after a revert

The `No Match` status is written to `MATCH_STATUS_COLUMN` and is not cleared when Layer 1 reverts the cell. Two distinct signals are being tracked:

- **Cell value** — reflects current state (restored to pre-edit value after revert)
- **Validation column** — reflects control history (attempt was made, mismatch was detected)

Clearing the flag on revert would make the sheet look clean when an unauthorized attempt had occurred. The `No Match` flag is a compliance signal, not a cell state indicator — it should persist.

### Why `No Owner Assigned` is an explicit outcome

Rows without a reviewer email in `OWNER_EMAIL_COLUMN` are treated as `Authorized` by default, consistent with how the enforcement layer handles unassigned rows. `No Owner Assigned` is logged explicitly rather than silently so that:

- The audit log reflects the full picture of review activity, including rows with missing assignment data
- Unassigned rows are visible and addressable — a pattern of `No Owner Assigned` entries is a signal that the review sheet setup has a gap

### Separation of enforcement and auditing

Enforcement and auditing are implemented as separate functions with no shared state. This is intentional:

- They are independent compliance concerns — one prevents incorrect completions, the other proves correct ones
- Each can be modified without touching the other — no regression risk across layers
- The audit function can be redirected to an external log destination by changing the append target in `onEditAuditHistory()` only, with no changes to enforcement logic

### Audit log extensibility

The audit log currently writes to a dedicated tab within the same Google Sheet. The decoupled architecture means the append target in `onEditAuditHistory()` is the only change required to redirect output to a centralized external audit store — for example, a shared audit sheet aggregating logs across all review sheets, or an external data destination. This was a deliberate design choice to keep the initial deployment simple while preserving the path to centralization.

### Reusability by configuration

All values that differ between deployments are isolated in two clearly marked configuration blocks at the top of the script. The logic blocks carry a `DO NOT MODIFY` marker. Deploying to a new review sheet requires only updating both configuration blocks and running through the deployment checklist below — no logic changes, no regression risk on existing deployments.

---

## Deployment Checklist

Required for every new sheet deployment. Steps must be completed in full — partial setup produces silent failures.

### 1. Set up an installable trigger

Simple triggers do not reliably capture `e.user`. An installable trigger is mandatory.

- Apps Script editor → Triggers (clock icon, left sidebar) → + Add Trigger
- Configure as follows:
  - Function to run: `onEdit`
  - Deployment: Head
  - Event source: From spreadsheet
  - Event type: On edit
- Complete the authorization flow
- Verify no duplicate simple `onEdit` trigger exists — delete it if present

**Note:** Triggers are bound to the spreadsheet, not the script. They do not carry over when a script is copied to a new sheet. This step must be repeated on every new deployment.

### 2. Create the Audit Log tab

- Add a blank tab named exactly as defined in `AUDIT_LOG_SHEET_NAME` — character for character, including spaces and capitalization
- Confirm the tab is **not** protected — the script must be able to append rows to it
- Add the following headers to row 1:

```
Timestamp | Editor Email | Record ID | User Name | Assigned Reviewer Email | New Value | Match Status | Edit Type
```

### 3. Add the validation column header

Add `Edit History Validation` to the cell at `MATCH_STATUS_COLUMN` in the header row of the main sheet.

### 4. Update both configuration blocks

**Layer 1 block:**

| Constant | What to set |
|---|---|
| `SHEET_NAME` | Exact name of the sheet tab to monitor |
| `COLUMNS_TO_WATCH` | Column numbers of the protected review decision fields |
| `OWNER_EMAIL_COLUMN` | Column number containing the assigned reviewer's email |
| `LOG_COLUMN_OFFSET` | Offset from edited column to reviewer timestamp column |

**Layer 2 block:**

| Constant | What to set |
|---|---|
| `AUDIT_LOG_SHEET_NAME` | Exact name of the audit log tab |
| `HISTORY_TRACK_COLUMN` | Column number of the primary decision field to audit |
| `MATCH_STATUS_COLUMN` | Column number for the Edit History Validation status |
| `AUDIT_RECORD_ID_COL` | Column number containing the system Record ID |
| `AUDIT_USER_NAME_COL` | Column number containing the User Name |

---

## Testing Matrix

| # | Action | Expected outcome |
|---|---|---|
| 1 | Assigned reviewer edits their row in the primary decision column | Edit accepted · `Match` written to validation column · Audit row logged as `Authorized` · Reviewer timestamp logged |
| 2 | Non-assigned user edits a row belonging to another reviewer | Edit reverted · Toast notification shown · `No Match` written to validation column · Audit row logged as `Unauthorized Attempt` |
| 3 | Row with no assigned reviewer — any user edits | Edit accepted · `No Owner Assigned` written to validation column · Audit row logged as `Authorized` |
| 4 | Multi-cell paste across owned and unowned rows | Each row evaluated independently — assigned rows accept, unassigned rows revert |
| 5 | Edit to a non-audited protected column (not `HISTORY_TRACK_COLUMN`) | Enforcement layer fires as normal · Validation column untouched · No audit log entry |
| 6 | Audit Log tab missing or renamed | Error written to Apps Script execution log · Main sheet unaffected · No crash |
| 7 | Two assigned reviewers editing simultaneously | Both edits accepted · Both rows receive `Match` · Two independent audit log rows appended |
| 8 | Script run manually from the Apps Script editor (no event object) | Both functions exit immediately on `!e` check · No changes to sheet |

---

## Common Failure Modes

| Symptom | Likely cause | Fix |
|---|---|---|
| Script runs but no reverts or logging | Installable trigger not set up | Set up installable trigger; confirm no duplicate simple trigger exists |
| Execution error on `appendRow()` | Audit Log tab is protected | Remove protection from the Audit Log tab |
| Script runs but does nothing | `SHEET_NAME` constant doesn't match the actual tab name exactly | Check for spacing or capitalization differences |
| `e.user` is null despite trigger being set | Authorization not completed when trigger was created | Delete trigger, re-create, complete the full authorization flow |
