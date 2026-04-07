/**
 * @OnlyCurrentDoc
 *
 * Access Review Controls Script — Row-Level Enforcement + Cell-Level Audit Layer
 *
 * Applicable to all access reviews — User Access Reviews (UARs) and
 * Privileged Access Reviews (PARs). Designed to be reused across any
 * access review sheet by updating the configuration blocks only.
 *
 * ─────────────────────────────────────────────────────────────────────────────
 * LAYER 1 — ROW-LEVEL APPROVER ENFORCEMENT (original)
 * ─────────────────────────────────────────────────────────────────────────────
 *   1. Protects COLUMNS_TO_WATCH — only the row's assigned reviewer may edit.
 *      The assigned reviewer is identified by their company Google account email,
 *      matched against the email stored in OWNER_EMAIL_COLUMN for that row.
 *   2. Unauthorized edits are reverted immediately and the user is notified
 *      via a toast pop-up.
 *   3. Authorized edits containing "yes" or "no" log a reviewer timestamp
 *      LOG_COLUMN_OFFSET columns to the right of the edited cell.
 *
 * ─────────────────────────────────────────────────────────────────────────────
 * LAYER 2 — CELL-LEVEL AUDIT LAYER (added)
 * ─────────────────────────────────────────────────────────────────────────────
 *   4. Every edit attempt to HISTORY_TRACK_COLUMN — authorized or not — is
 *      appended as a new row to the dedicated audit log tab defined in
 *      AUDIT_LOG_SHEET_NAME. Fields logged:
 *        Timestamp | Editor Email | Record ID | User Name | Assigned Reviewer
 *        Email | New Value | Match Status | Edit Type
 *   5. All edit outcomes write a validation status to MATCH_STATUS_COLUMN
 *      in the main sheet — "Match", "No Match", or "No Owner Assigned" —
 *      surfacing approver mismatches visually for at-a-glance compliance review.
 *   6. Unauthorized edit attempts are logged as "Unauthorized Attempt" in the
 *      audit log with the full event detail captured before the revert fires,
 *      ensuring the compliance record is complete even when the control succeeds.
 *
 * Design note: The audit log is maintained as a dedicated tab within the same
 * Google Sheet. The architecture is intentionally decoupled — the logging
 * function is independent of the enforcement function — and can be extended
 * to write to an external audit destination (e.g., a centralized log sheet or
 * external store) by modifying the append target in onEditAuditHistory() only.
 *
 * ─────────────────────────────────────────────────────────────────────────────
 * PRE-REQUISITES — required for every new sheet deployment
 * ─────────────────────────────────────────────────────────────────────────────
 *
 *   1. INSTALLABLE TRIGGER (critical — script will not work without this)
 *      Simple triggers do not reliably capture e.user for all editors in a
 *      Google Workspace environment. An installable trigger is mandatory.
 *      - Open this sheet's Apps Script editor
 *      - Go to Triggers (clock icon, left sidebar)
 *      - Click + Add Trigger and configure as follows:
 *          Function to run : onEdit
 *          Deployment      : Head
 *          Event source    : From spreadsheet
 *          Event type      : On edit
 *      - Save and complete the authorization flow
 *      - Verify no duplicate simple onEdit trigger exists — delete it if so
 *      Note: triggers are bound to the spreadsheet, not the script. They do
 *      NOT carry over when the script is copied to a new sheet. This step
 *      must be repeated manually on every new sheet deployment.
 *
 *   2. AUDIT LOG TAB
 *      - Create a blank tab named exactly as defined in AUDIT_LOG_SHEET_NAME
 *      - Ensure the tab is NOT protected — the script must be able to append rows
 *      - Add the following headers to row 1:
 *        Timestamp | Editor Email | Record ID | User Name | Assigned Reviewer
 *        Email | New Value | Match Status | Edit Type
 *
 *   3. MAIN SHEET HEADER
 *      - Add the header "Edit History Validation" to the cell at
 *        MATCH_STATUS_COLUMN in the header row of the main sheet
 *
 *   4. UPDATE CONFIGURATION BLOCKS
 *      - Update all values in both configuration blocks below to match the
 *        target sheet. Only the configuration blocks should ever change
 *        between deployments — the logic blocks are not modified.
 */


// =============================================================================
// LAYER 1 CONFIGURATION — update these values to match your sheet
// =============================================================================
const SHEET_NAME         = "Access Review";  // Exact name of the sheet tab to monitor
const COLUMNS_TO_WATCH   = [25, 26, 27];     // Columns to enforce (Y=25, Z=26, AA=27)
const OWNER_EMAIL_COLUMN = 23;               // Column containing the assigned reviewer's email (W=23)
const LOG_COLUMN_OFFSET  = 4;                // Reviewer timestamp: edited col + offset (e.g. Y+4 → col AC)
// =============================================================================


// =============================================================================
// LAYER 1 — ROW-LEVEL ENFORCEMENT
// Do not modify this function except through the configuration block above.
// =============================================================================
function onEdit(e) {
  onEditAuditHistory(e); // ← Passes every edit to the audit layer first, before any revert fires

  if (!e || !e.range) {
    return;
  }

  const range     = e.range;
  const sheet     = range.getSheet();
  const userEmail = e.user ? e.user.getEmail() : null;

  if (sheet.getName() !== SHEET_NAME || !userEmail) {
    return;
  }

  for (let row = range.getRow(); row <= range.getLastRow(); row++) {
    for (let col = range.getColumn(); col <= range.getLastColumn(); col++) {

      if (COLUMNS_TO_WATCH.includes(col)) {
        const ownerEmail = sheet.getRange(row, OWNER_EMAIL_COLUMN).getValue().trim().toLowerCase();

        if (ownerEmail && userEmail.toLowerCase() !== ownerEmail) {

          // Revert unauthorized edit.
          // For single-cell edits, restore e.oldValue.
          // For multi-cell edits (drag/paste), e.oldValue is undefined — revert to blank as safe fallback.
          const oldValue = (range.getNumRows() === 1 && range.getNumColumns() === 1) ? e.oldValue : "";
          sheet.getRange(row, col).setValue(oldValue);

          SpreadsheetApp.getActiveSpreadsheet().toast(
            `Change reverted. You are not authorized to edit this row.`,
            "Access Denied", 5
          );

        } else {

          // Log reviewer timestamp on authorized edits containing "yes" or "no"
          const cell     = sheet.getRange(row, col);
          const newValue = cell.getValue().toString().toLowerCase();

          if (newValue.includes("yes") || newValue.includes("no")) {
            const timestamp = new Date();
            const stampText = `${userEmail} reviewed on ${timestamp.toLocaleString()}`;
            sheet.getRange(row, col + LOG_COLUMN_OFFSET).setValue(stampText);
          }
        }
      }
    }
  }
}
// =============================================================================


// =============================================================================
// LAYER 2 CONFIGURATION — update these values to match your sheet
// All values that need to change per deployment are defined here.
// Do not modify anything below this configuration block.
// <<<<<<<<<<<<<<<<<<<<< UPDATE THESE VALUES TO MATCH YOUR SHEET >>>>>>>>>>>>>>>>>

const AUDIT_LOG_SHEET_NAME = "Audit Log";  // Exact name of the audit log tab
const HISTORY_TRACK_COLUMN = 25;           // Primary decision column to audit (Y=25) — e.g. "Approve current access?"
const MATCH_STATUS_COLUMN  = 30;           // Column for Edit History Validation status in main sheet (AD=30)
const AUDIT_RECORD_ID_COL  = 1;            // Column containing the system Record ID (A=1)
const AUDIT_USER_NAME_COL  = 2;            // Column containing the User Name (B=2)

// Note: SHEET_NAME and OWNER_EMAIL_COLUMN are shared with the Layer 1 configuration
// block above. Update them there when deploying to a new sheet.

// <<<<<<<<<<<<<<<<<<<<< DO NOT MODIFY ANYTHING BELOW THIS LINE >>>>>>>>>>>>>>>>>
// =============================================================================


/**
 * Layer 2 — Cell-level audit handler for HISTORY_TRACK_COLUMN.
 *
 * Called as the first action inside onEdit() on every edit event — before
 * the Layer 1 revert logic runs — so that unauthorized values are captured
 * in the audit log accurately before they are wiped.
 *
 * Flow per cell touched in HISTORY_TRACK_COLUMN:
 *   1. Reads editor email, assigned reviewer email, row metadata, and new cell value.
 *   2. Compares editor email vs. assigned reviewer email → derives Match Status and Edit Type.
 *   3. Appends a full record to the audit log sheet regardless of authorization outcome.
 *   4. Writes Edit History Validation status to MATCH_STATUS_COLUMN in the main sheet
 *      for all edit outcomes — including unauthorized attempts.
 *
 * Three possible outcomes:
 *   Match             — editor is the assigned reviewer for that row (Authorized)
 *   No Match          — editor is not the assigned reviewer (Unauthorized Attempt)
 *   No Owner Assigned — no reviewer email in the row (Authorized by default)
 *
 * "No Match" is intentionally written to MATCH_STATUS_COLUMN even after the cell
 * is reverted by Layer 1 — the mismatch attempt is itself a compliance signal and
 * must be surfaced, not just logged in the audit trail.
 *
 * This function does not revert unauthorized edits. That is handled exclusively
 * by the Layer 1 onEdit() function above. Separation of responsibilities is
 * intentional — enforcement and auditing are independent concerns.
 */
function onEditAuditHistory(e) {

  if (!e || !e.range) return;

  const range     = e.range;
  const sheet     = range.getSheet();
  const userEmail = e.user ? e.user.getEmail().toLowerCase() : null;

  if (sheet.getName() !== SHEET_NAME || !userEmail) return;

  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const auditSheet = ss.getSheetByName(AUDIT_LOG_SHEET_NAME);

  if (!auditSheet) {
    console.error(
      `Audit log sheet "${AUDIT_LOG_SHEET_NAME}" not found. ` +
      `Create a blank tab with that exact name before deploying.`
    );
    return;
  }

  for (let row = range.getRow(); row <= range.getLastRow(); row++) {
    for (let col = range.getColumn(); col <= range.getLastColumn(); col++) {

      if (col !== HISTORY_TRACK_COLUMN) continue;

      const recordId      = sheet.getRange(row, AUDIT_RECORD_ID_COL).getValue();
      const userName      = sheet.getRange(row, AUDIT_USER_NAME_COL).getValue();
      const reviewerEmail = sheet.getRange(row, OWNER_EMAIL_COLUMN).getValue().trim().toLowerCase();
      const newValue      = sheet.getRange(row, col).getValue();
      const timestamp     = new Date();

      let matchStatus, editType;

      if (!reviewerEmail) {
        matchStatus = "No Owner Assigned";
        editType    = "Authorized";
      } else if (userEmail === reviewerEmail) {
        matchStatus = "Match";
        editType    = "Authorized";
      } else {
        matchStatus = "No Match";
        editType    = "Unauthorized Attempt";
      }

      // Append to audit log
      // Column order matches the required header row in the audit tab:
      // Timestamp | Editor Email | Record ID | User Name | Assigned Reviewer Email |
      // New Value | Match Status | Edit Type
      auditSheet.appendRow([
        timestamp,
        userEmail,
        recordId,
        userName,
        reviewerEmail,
        newValue,
        matchStatus,
        editType
      ]);

      // Write validation status to main sheet — all outcomes including unauthorized attempts
      sheet.getRange(row, MATCH_STATUS_COLUMN).setValue(matchStatus);
    }
  }
}
