/**
 * ============================================
 * OT BILLING SYSTEM - GOOGLE APPS SCRIPT  (v45-compatible)
 * ============================================
 *
 * SETUP INSTRUCTIONS (for Markus):
 * 1. Upload the OT_Billing_System .xlsx to Google Drive.
 * 2. Open it, then File > Save as Google Sheets (must be a native Google Sheet).
 * 3. In the Google Sheet: Extensions > Apps Script.
 * 4. Delete any existing code, paste this entire script, and click Save.
 * 5. In the script editor, run the function `installTriggers` once.
 *    - Google will ask you to authorize. Choose your account,
 *      click "Advanced" > "Go to project (unsafe)" > Allow.
 * 6. Go back to the spreadsheet and refresh the page.
 *    You'll see a new "📧 Invoice System" menu.
 * 7. Business info (OT WITH MJ, PLLC letterhead) is already filled in
 *    under CONFIG.BUSINESS below. Optionally add a phone number and/or
 *    email there - if email is left blank, your Gmail address is used.
 *
 * HOW TO FIND A FILE/FOLDER ID:
 * Open the file in Google Drive. The URL looks like:
 *   https://docs.google.com/spreadsheets/d/THIS_LONG_STRING/edit
 * Copy the long string between /d/ and /edit.
 */

// ==========================================
// CONFIGURATION
// ==========================================

const CONFIG = {
  // Templates are found automatically by NAME in Google Drive - no IDs needed.
  // Upload the template files and convert them to Google formats:
  //   - invoice_template.xlsx -> open it > File > Save as Google Sheets
  //   - email_template.docx   -> open it > File > Save as Google Docs
  // The search matches the name with or without the file extension.
  // If a template isn't found, the script falls back to a built-in layout/text.
  INVOICE_TEMPLATE_NAME: 'invoice_template',
  EMAIL_TEMPLATE_NAME: 'email_template',

  // OPTIONAL: pin templates to exact files by ID (overrides the name search).
  // Useful if there are multiple files with the same name.
  INVOICE_TEMPLATE_ID: '',
  EMAIL_TEMPLATE_ID: '',

  // PDFs are saved automatically to a subfolder named "invoices" created
  // next to this spreadsheet in Drive (no ID needed). Set an ID below only
  // if you want to force a specific existing folder instead.
  INVOICES_FOLDER_ID: '',
  INVOICES_FOLDER_NAME: 'invoices',

  // Business information (from OT WITH MJ letterhead)
  BUSINESS: {
    name: 'OT WITH MJ, PLLC.',
    practitioner: 'Markus Jarrow, OTR/L, C/NDT',
    address: '10 Grand Avenue #37',
    cityState: 'Brooklyn, NY 11205',
    license: '009131',       // NY License #
    npi: '1497900062',       // NPI #
    tin: '135-62-8603',      // TIN #
    phone: '',               // optional - leave '' to omit from invoices/emails
    email: ''                // optional - defaults to the Gmail account running the script
  },

  // Invoice settings
  PAYMENT_TERMS_DAYS: 30,
  INVOICE_PREFIX: '',          // e.g. 'INV-' ; sheet currently uses plain numbers (1507, 1508, ...)
  DEFAULT_RATE_PER_UNIT: 111,  // used only if Units is blank/zero

  // ---- 💰 Invoices sheet layout (1-based column indexes) ----
  // A: (blank)  B: 💵  C: Child's Name  D: Start Date  E: End Date
  // F: Created On  G: Units  H: Hours  I: Total Amount  J: Invoice #
  // K: Email Sent?  L: Zelle?  M: Amount Paid  N: Send Invoice
  COLUMNS: {
    CHILD_NAME: 3,    // C
    PERIOD_START: 4,  // D
    PERIOD_END: 5,    // E
    CREATED_ON: 6,    // F
    UNITS: 7,         // G
    HOURS: 8,         // H
    TOTAL: 9,         // I
    INVOICE_NUM: 10,  // J
    EMAIL_SENT: 11,   // K  (✓ / ✗)
    ZELLE: 12,        // L  (✓ / ✗)
    AMOUNT_PAID: 13,  // M
    SEND_INVOICE: 14  // N  (check to send)
  },
  DATA_START_ROW: 2,  // row 1 is the header

  // ---- 👥 Clients sheet layout (1-based column indexes) ----
  // A: Child's Name  B: Client Name (parent)  C: DOB  D: Diagnosis #
  // E: Phone  F: Street Address  G: City  H: State  I: Zipcode  J: Email
  CLIENT_COLUMNS: {
    CHILD_NAME: 1,   // A
    PARENT_NAME: 2,  // B
    DOB: 3,          // C
    DIAGNOSIS: 4,    // D
    PHONE: 5,        // E
    STREET: 6,       // F
    CITY: 7,         // G
    STATE: 8,        // H
    ZIP: 9,          // I
    EMAIL: 10        // J  (may contain several addresses separated by commas)
  },

  // Sheet (tab) names - must match exactly, including the emoji
  SHEETS: {
    INVOICES: '💰 Invoices',
    CLIENTS: '👥 Clients',
    SESSION_DATA: 'Session Data'
  },

  // Values printed in the Services Delivered table for columns the
  // workbook doesn't track per-session (matches past invoices)
  SERVICE_LOCATION: 'Home',
  CPT_MODIFIER: 'GO',

  // Status marks used in the sheet
  MARKS: { YES: '✓', NO: '✗' }
};

// ==========================================
// MENU AND TRIGGERS
// ==========================================

/**
 * Creates the custom menu when the spreadsheet opens.
 * (Simple trigger - runs automatically, no installation needed.)
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📧 Invoice System')
    .addItem('📧 Send Invoice for Selected Row', 'sendInvoiceForSelectedRow')
    .addSeparator()
    .addItem('✅ Mark Selected as Paid in Full', 'markSelectedAsPaid')
    .addItem('🔄 Clear Payment on Selected', 'markSelectedAsUnpaid')
    .addSeparator()
    .addItem('🔍 Check Gmail for Zelle Payments', 'checkForZellePayments')
    .addSeparator()
    .addItem('🧾 Generate Invoices', 'generateInvoicesFromSessions')
    .addToUi();

  createMissingClientTabs();
}

/**
 * Installable edit trigger - created by installTriggers().
 * When the "Send Invoice" checkbox (column N) is checked on the
 * 💰 Invoices sheet, the invoice for that row is emailed automatically.
 *
 * NOTE: this is deliberately NOT named onEdit. A simple onEdit trigger
 * runs without authorization and cannot send email; and having both a
 * simple and an installable trigger with the same name makes every
 * edit fire twice.
 */
function handleEditInstallable(e) {
  if (!e || !e.range) return; // guard against running it manually from the editor

  const sheet = e.range.getSheet();
  if (sheet.getName() !== CONFIG.SHEETS.INVOICES) return;

  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row < CONFIG.DATA_START_ROW) return;
  if (col !== CONFIG.COLUMNS.SEND_INVOICE) return;

  const value = e.range.getValue();
  const isChecked = value === true || value === CONFIG.MARKS.YES || value === 'TRUE';
  if (!isChecked) return;

  try {
    processInvoiceRow(sheet, row); // marks "Email Sent?" itself
    e.range.setValue(false); // reset the checkbox
  } catch (error) {
    e.range.setValue(false);
    SpreadsheetApp.getActiveSpreadsheet().toast('❌ ' + error.message, 'Invoice not sent', 10);
  }
}

/**
 * Run this ONCE from the script editor to set up the triggers.
 */
function installTriggers() {
  // Remove only this project's previous triggers to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('handleEditInstallable')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();

  SpreadsheetApp.getUi().alert(
    '✅ Triggers installed.\n\n' +
    'Checking a box in the "Send Invoice" column (N) on the 💰 Invoices ' +
    'sheet will now email that invoice automatically.'
  );
}

/**
 * "Generate Invoices for New Sessions" menu action.
 *
 * For every 👶 child tab, finds session rows that have data (a date and
 * units) but haven't been billed yet, sums them per child into a single
 * new invoice, inserts that invoice at the TOP of the 💰 Invoices table
 * (row 2, just under the header), and marks those session rows as
 * invoiced so re-running this never double-bills them.
 *
 * "Already invoiced" is tracked in column H of each 👶 tab (added next
 * to the existing columns) - holds the Invoices row's Invoice # once
 * billed, blank otherwise. This survives even if Invoices rows are
 * later reordered or deleted.
 */
function generateInvoicesFromSessions() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invoicesSheet = ss.getSheetByName(CONFIG.SHEETS.INVOICES);
  if (!invoicesSheet) {
    ui.alert('Sheet "' + CONFIG.SHEETS.INVOICES + '" was not found.');
    return;
  }

  const childTabs = ss.getSheets().filter(sh => sh.getName().indexOf('👶 ') === 0);
  const summaries = [];

  for (const sheet of childTabs) {
    const childName = sheet.getName().replace('👶 ', '').trim();
    const unbilled = getUnbilledSessions(sheet);
    if (unbilled.rows.length === 0) continue;
    summaries.push({ sheet, childName, ...unbilled });
  }

  if (summaries.length === 0) {
    ui.alert('No new (un-invoiced) sessions were found in any client tab.');
    return;
  }

  // Insert newest-first: process in reverse so the very first child
  // summarized ends up at the very top (row 2) after all are inserted.
  const createdLabels = [];
  for (let i = summaries.length - 1; i >= 0; i--) {
    const s = summaries[i];
    const invoiceNumber = insertInvoiceRow(invoicesSheet, s);
    markSessionsAsInvoiced(s.sheet, s.rowNumbers, invoiceNumber);
    createdLabels.push(s.childName + ' (#' + invoiceNumber + ', ' +
      formatCurrency(s.total) + ', ' + s.rowNumbers.length + ' session(s))');
  }

  SpreadsheetApp.flush();
  ui.alert('✅ Created ' + createdLabels.length + ' invoice(s):\n\n' +
    createdLabels.join('\n'));
}

/**
 * Reads a 👶 tab and returns its un-invoiced sessions: rows with a date
 * and units but no value yet in column H ("Invoiced As"). Returns
 * { rows: [...], rowNumbers: [...], units, hours, total,
 *   periodStart, periodEnd } - hours/total mirror the column G formula
 * convention (units * rate; "hours" here follows the workbook's existing
 * units-as-half-hours pattern used elsewhere in this script).
 */
function getUnbilledSessions(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { rows: [], rowNumbers: [], units: 0, hours: 0, total: 0 };

  // Columns: A name, B date, C/D CPT codes, E units, F rate, G total, H invoiced-as
  const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  const rows = [];
  const rowNumbers = [];
  let units = 0, total = 0;
  let minDate = null, maxDate = null;

  data.forEach((row, i) => {
    const [, date, , , unitsVal, , totalVal, invoicedAs] = row;
    if (!(date instanceof Date)) return;          // blank/template row
    if (Number(unitsVal) <= 0 || unitsVal === '') return;
    if (invoicedAs) return;                        // already billed

    rows.push(row);
    rowNumbers.push(i + 2); // sheet row number
    units += Number(unitsVal) || 0;
    total += Number(totalVal) || 0;
    if (!minDate || date < minDate) minDate = date;
    if (!maxDate || date > maxDate) maxDate = date;
  });

  // Workbook convention elsewhere: hours = units / 2 (e.g. 14 units = 7
  // hours on Sofia's existing invoice). Kept consistent here.
  return {
    rows, rowNumbers, units, total,
    hours: units / 2,
    periodStart: minDate, periodEnd: maxDate
  };
}

/**
 * Inserts a new row at the top of the Invoices table (row 2, under the
 * header) for one child's batch of new sessions, matching the existing
 * row formatting and the "Send Invoice" checkbox widget. Returns the
 * assigned invoice number.
 */
function insertInvoiceRow(sheet, summary) {
  const C = CONFIG.COLUMNS;
  sheet.insertRowAfter(1);

  // Copy formatting from what is now row 3 (the table's previous top
  // data row, or the header if the table was empty) so the new row
  // matches the sheet's existing look exactly.
  const formatSourceRow = sheet.getLastRow() >= 3 ? 3 : 1;
  sheet.getRange(formatSourceRow, 1, 1, sheet.getLastColumn())
    .copyTo(sheet.getRange(2, 1, 1, sheet.getLastColumn()),
      SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

  sheet.getRange(2, C.CHILD_NAME).setValue(summary.childName);
  sheet.getRange(2, C.PERIOD_START).setValue(summary.periodStart);
  sheet.getRange(2, C.PERIOD_END).setValue(summary.periodEnd);
  sheet.getRange(2, C.CREATED_ON).setValue(new Date());
  sheet.getRange(2, C.UNITS).setValue(summary.units);
  sheet.getRange(2, C.HOURS).setValue(summary.hours);
  sheet.getRange(2, C.TOTAL).setValue(summary.total);
  sheet.getRange(2, 2).setValue('💵'); // column B icon, matches existing rows
  sheet.getRange(2, C.EMAIL_SENT).setValue(CONFIG.MARKS.NO)
    .setFontColor('#FF0000').setHorizontalAlignment('center');
  sheet.getRange(2, C.ZELLE).setValue(CONFIG.MARKS.NO)
    .setFontColor('#FF0000').setHorizontalAlignment('center');
  sheet.getRange(2, C.SEND_INVOICE).insertCheckboxes().setValue(false);

  const invoiceNumber = getOrAssignInvoiceNumber(sheet, 2);
  // Column J always holds the plain number; use that same value (not the
  // prefixed display string) when stamping the child tab in column H,
  // so the two stay comparable.
  return sheet.getRange(2, C.INVOICE_NUM).getValue();
}

/**
 * Stamps the given session rows on a child's tab with the invoice
 * number that just billed them, in column H ("Invoiced As"). Adds the
 * header to column H if it isn't there yet (older tabs predate it).
 */
function markSessionsAsInvoiced(sheet, rowNumbers, invoiceNumber) {
  if (sheet.getRange(1, 8).getValue() !== 'Invoiced As') {
    sheet.getRange(1, 8).setValue('Invoiced As')
      .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#FF69B4');
  }
  for (const row of rowNumbers) {
    sheet.getRange(row, 8).setValue(invoiceNumber);
  }
}

/**
 * Scans the 👥 Clients sheet for any child without a matching '👶 <name>'
 * tab and creates one, copying the layout (header row, name pre-filled
 * down column A, =E*F formula down column G) from an existing 👶 tab.
 * Runs automatically every time the spreadsheet is opened (called from
 * onOpen) - no menu click needed.
 *
 * Safe to run repeatedly: a child that already has a tab is skipped.
 */
function createMissingClientTabs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const clientsSheet = ss.getSheetByName(CONFIG.SHEETS.CLIENTS);
  if (!clientsSheet) return; // nothing to scan against

  const K = CONFIG.CLIENT_COLUMNS;
  const data = clientsSheet.getDataRange().getValues();
  const childNames = [];
  for (let i = 1; i < data.length; i++) { // skip header
    const name = String(data[i][K.CHILD_NAME - 1] || '').trim();
    if (name) childNames.push(name);
  }
  if (childNames.length === 0) return;

  // Find an existing 👶 tab to use as the template for layout/formulas.
  // Falls back to a hardcoded layout if this is the very first client tab.
  const templateSheet = ss.getSheets().find(sh => sh.getName().indexOf('👶 ') === 0);

  let created = [];
  for (const name of childNames) {
    const tabName = '👶 ' + name;
    if (ss.getSheetByName(tabName)) continue; // already exists

    const newSheet = templateSheet ?
      templateSheet.copyTo(ss) :
      ss.insertSheet();
    newSheet.setName(tabName);

    if (templateSheet) {
      // copyTo duplicates EVERYTHING from the template, including its
      // actual session data (dates, CPT codes, units, rates) - clear all
      // data columns before writing the new child's name, so no other
      // client's billing history rides along into the new tab.
      const lastRow = Math.max(newSheet.getLastRow(), 2);
      const lastCol = Math.max(newSheet.getLastColumn(), 7);
      newSheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
      newSheet.getRange(2, 1, lastRow - 1, 1).setValue(name);
      // Re-apply the G column formula (clearContent removed it too)
      const formulas = [];
      for (let r = 2; r <= lastRow; r++) formulas.push(['=E' + r + '*F' + r]);
      newSheet.getRange(2, 7, lastRow - 1, 1).setFormulas(formulas);
    } else {
      buildBlankClientTab(newSheet, name);
    }
    created.push(name);
  }

  // Move each new tab to sit right after the Clients tab, in Clients order
  if (created.length > 0) {
    for (const name of childNames) {
      const sh = ss.getSheetByName('👶 ' + name);
      if (sh && created.indexOf(name) !== -1) {
        ss.setActiveSheet(sh);
        ss.moveActiveSheet(clientsSheet.getIndex() + 1);
      }
    }
    ss.toast('Created tab(s) for: ' + created.join(', '), '📧 Invoice System', 6);
  }
}

/**
 * Builds a 👶 tab from scratch, matching the standard client tab style
 * exactly (pink header, white bold text, column widths, currency/date
 * formatting). Used whenever no existing 👶 tab is found to copy as a
 * template - including after every client tab has been deleted, so the
 * style is never lost even with zero clients in the workbook.
 */
function buildBlankClientTab(sheet, childName) {
  const headers = ["Child's Name", 'Date', 'CPT Code #1', 'CPT Code #2',
    'Units', 'Rate per Unit', 'Total Amount'];
  const rows = 200;

  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setFontColor('#FFFFFF')
    .setFontFamily('Calibri')
    .setBackground('#FF69B4');

  sheet.getRange(2, 1, rows, 1).setValue(childName);

  const formulas = [];
  for (let r = 2; r <= rows + 1; r++) formulas.push(['=E' + r + '*F' + r]);
  sheet.getRange(2, 7, rows, 1).setFormulas(formulas);

  sheet.getRange(2, 2, rows, 1).setNumberFormat('M/d/yyyy');   // Date
  sheet.getRange(2, 6, rows, 1).setNumberFormat('$#,##0.00');  // Rate per Unit
  sheet.getRange(2, 7, rows, 1).setNumberFormat('$#,##0.00');  // Total Amount

  sheet.setColumnWidth(1, 21.57 * 7); // pt-per-char approximation for A:G
  sheet.setColumnWidth(2, 14.43 * 7);
  sheet.setColumnWidth(5, 8.71 * 7);
  sheet.setColumnWidth(6, 14.43 * 7);
  sheet.setColumnWidth(7, 15.86 * 7);
  sheet.setFrozenRows(1);
}

// ==========================================
// PAYMENT FUNCTIONS
// ==========================================

/**
 * Marks the selected invoice row as paid in full:
 * copies Total Amount (I) into Amount Paid (M).
 */
function markSelectedAsPaid() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();

  if (sheet.getName() !== CONFIG.SHEETS.INVOICES) {
    ui.alert('Please select a row on the "' + CONFIG.SHEETS.INVOICES + '" sheet.');
    return;
  }
  const row = sheet.getActiveRange().getRow();
  if (row < CONFIG.DATA_START_ROW) {
    ui.alert('Please select a data row (row ' + CONFIG.DATA_START_ROW + ' or below).');
    return;
  }

  const total = sheet.getRange(row, CONFIG.COLUMNS.TOTAL).getValue();
  sheet.getRange(row, CONFIG.COLUMNS.AMOUNT_PAID).setValue(total);
  ui.alert('✅ Amount Paid set to ' + formatCurrency(total));
}

/**
 * Clears the payment on the selected invoice row.
 */
function markSelectedAsUnpaid() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();

  if (sheet.getName() !== CONFIG.SHEETS.INVOICES) {
    ui.alert('Please select a row on the "' + CONFIG.SHEETS.INVOICES + '" sheet.');
    return;
  }
  const row = sheet.getActiveRange().getRow();
  if (row < CONFIG.DATA_START_ROW) {
    ui.alert('Please select a data row (row ' + CONFIG.DATA_START_ROW + ' or below).');
    return;
  }

  sheet.getRange(row, CONFIG.COLUMNS.AMOUNT_PAID).clearContent();
  sheet.getRange(row, CONFIG.COLUMNS.ZELLE)
    .setValue(CONFIG.MARKS.NO)
    .setFontColor('#FF0000')
    .setHorizontalAlignment('center');
  ui.alert('✅ Payment cleared.');
}

/**
 * Scans Gmail for Zelle payment notifications matching unpaid invoices.
 *
 * For every invoice row that isn't fully paid (Amount Paid < Total), it
 * searches the mailbox for Zelle emails mentioning the parent's name and
 * checks that the invoice amount appears in the message. On a match:
 * Zelle? -> ✓ (green), Amount Paid -> Total Amount. A summary alert
 * lists what was found and what wasn't.
 *
 * Searches the last ZELLE_SEARCH_DAYS days (see below).
 */
function checkForZellePayments() {
  const ZELLE_SEARCH_DAYS = 30;
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEETS.INVOICES);
  if (!sheet) {
    ui.alert('Sheet "' + CONFIG.SHEETS.INVOICES + '" was not found.');
    return;
  }

  const C = CONFIG.COLUMNS;
  const lastRow = sheet.getLastRow();
  const found = [];
  const notFound = [];
  const usedMessageIds = new Set(); // each payment email can pay ONE invoice
  const windowStart = new Date();
  windowStart.setDate(windowStart.getDate() - ZELLE_SEARCH_DAYS);

  for (let row = CONFIG.DATA_START_ROW; row <= lastRow; row++) {
    const childName = sheet.getRange(row, C.CHILD_NAME).getValue();
    if (!childName) continue;

    const total = Number(sheet.getRange(row, C.TOTAL).getValue()) || 0;
    const amountPaid = Number(sheet.getRange(row, C.AMOUNT_PAID).getValue()) || 0;
    if (!total || amountPaid >= total) continue; // already paid (or no total)

    const invoiceNum = sheet.getRange(row, C.INVOICE_NUM).getValue();
    const client = getClientInfo(childName);
    const parentName = client ? String(client.parentName || '').trim() : '';
    const label = childName + (invoiceNum ? ' (#' + invoiceNum + ')' : '');

    if (!parentName) {
      notFound.push(label + ' - no parent name on file');
      continue;
    }

    // A payment for THIS invoice can only have arrived after the invoice
    // was created - so ignore anything older than Created On (column F),
    // capped at the 30-day window. This stops an old Zelle payment of the
    // same recurring amount from marking a new invoice as paid.
    const createdOn = sheet.getRange(row, C.CREATED_ON).getValue();
    const afterDate = (createdOn instanceof Date && createdOn > windowStart) ?
      createdOn : windowStart;

    if (findZellePayment(parentName, total, afterDate, usedMessageIds)) {
      sheet.getRange(row, C.ZELLE)
        .setValue(CONFIG.MARKS.YES)
        .setFontColor('#188038')
        .setHorizontalAlignment('center');
      sheet.getRange(row, C.AMOUNT_PAID).setValue(total);
      found.push(label + ' - ' + formatCurrency(total) + ' from ' + parentName);
    } else {
      notFound.push(label + ' - ' + formatCurrency(total));
    }
  }

  SpreadsheetApp.flush();

  if (found.length === 0 && notFound.length === 0) {
    ui.alert('✅ All invoices are already marked as paid.');
    return;
  }
  let message = '';
  if (found.length > 0) {
    message += '✅ Zelle payments found and marked paid:\n' +
      found.join('\n') + '\n\n';
  }
  if (notFound.length > 0) {
    message += '❓ No Zelle payment found for:\n' + notFound.join('\n');
  }
  ui.alert(message.trim());
}

/**
 * Searches Gmail for a Zelle notification from `senderName` containing
 * the invoice amount. Returns true if a matching message is found.
 *
 * Matching is on the CLIENT (parent) name, approximately:
 * first and last name are matched independently and case-insensitively,
 * so "TARYN BURNS" matches "Taryn E. Burns sent you $222.00".
 *
 * Guards against false positives from recurring identical amounts:
 *   - only messages dated AFTER `afterDate` count (the invoice's
 *     Created On date, capped at the 30-day window) - an email that
 *     predates the invoice can't be a payment for it
 *   - each Gmail message can satisfy only ONE invoice per run
 *     (`usedMessageIds`), so one payment can't mark two invoices paid
 *
 * A message counts as a payment only if it has all of:
 *   1. the word "zelle"
 *   2. the parent's first AND last name (any case, middle names ignored)
 *   3. the invoice amount ($1,554.00 / $1554.00 / $1,554 formats)
 *   4. a date on/after the invoice's Created On date
 */
function findZellePayment(senderName, amount, afterDate, usedMessageIds) {
  const nameParts = String(senderName).toLowerCase().split(/\s+/)
    .filter(p => p.replace(/[^a-z]/g, '').length > 1); // drop initials/empties
  if (nameParts.length === 0) return false;
  const first = nameParts[0];
  const last = nameParts[nameParts.length - 1];

  const amountStr = formatCurrency(amount);                    // $1,554.00
  const amountPlain = amountStr.replace(/,/g, '');             // $1554.00
  const amountNoCents = amountStr.replace(/\.00$/, '');        // $1,554

  // Gmail's after: operator is day-granular; the exact timestamp check
  // happens per-message below.
  const afterStr = Utilities.formatDate(afterDate,
    Session.getScriptTimeZone(), 'yyyy/MM/dd');
  const query = 'zelle "' + first + '"' +
    (last !== first ? ' "' + last + '"' : '') +
    ' after:' + afterStr;
  const threads = GmailApp.search(query, 0, 20);

  for (const thread of threads) {
    for (const message of thread.getMessages()) {
      if (usedMessageIds.has(message.getId())) continue;
      if (message.getDate() < afterDate) continue;

      const text = (message.getSubject() + '\n' + message.getPlainBody());
      const lower = text.toLowerCase();
      const nameOk = lower.indexOf(first) !== -1 &&
                     lower.indexOf(last) !== -1;
      const amountOk = text.indexOf(amountStr) !== -1 ||
                       text.indexOf(amountPlain) !== -1 ||
                       text.indexOf(amountNoCents) !== -1;
      if (nameOk && amountOk) {
        usedMessageIds.add(message.getId()); // consume: one email = one invoice
        return true;
      }
    }
  }
  return false;
}

// ==========================================
// MAIN INVOICE FUNCTIONS
// ==========================================

/**
 * Menu action: builds the invoice for the selected row and shows a
 * preview dialog - View PDF / Send / Cancel - before anything is emailed.
 */
function sendInvoiceForSelectedRow() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();

  if (sheet.getName() !== CONFIG.SHEETS.INVOICES) {
    ui.alert('Please select a row on the "' + CONFIG.SHEETS.INVOICES + '" sheet.');
    return;
  }
  const row = sheet.getActiveRange().getRow();
  if (row < CONFIG.DATA_START_ROW) {
    ui.alert('Please select a data row (row ' + CONFIG.DATA_START_ROW + ' or below).');
    return;
  }

  let draft;
  try {
    draft = prepareInvoiceDraft(sheet, row);
  } catch (error) {
    ui.alert('❌ Error: ' + error.message);
    console.error(error);
    return;
  }

  const esc = s => String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');

  const warningRow = draft.warning ?
    '<tr><td class="l">⚠️</td><td style="color:#b45309;">' + esc(draft.warning) + '</td></tr>' : '';

  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 4px 10px; color: #222; }
      table { border-collapse: collapse; margin: 8px 0 14px; }
      td { padding: 3px 10px 3px 0; vertical-align: top; }
      td.l { color: #666; white-space: nowrap; }
      .btns { margin-top: 10px; }
      button, a.view {
        font-size: 13px; padding: 8px 16px; margin-right: 8px;
        border-radius: 4px; border: 1px solid #ccc; background: #f5f5f5;
        cursor: pointer; text-decoration: none; color: #222; display: inline-block;
      }
      button.send { background: #2E7D32; border-color: #2E7D32; color: #fff; }
      #status { margin-top: 10px; color: #666; }
    </style>
    <h3 style="margin:6px 0;">Invoice #${esc(draft.invoiceNumber)} - ready to send</h3>
    <table>
      <tr><td class="l">Client:</td><td>${esc(draft.childName)} (${esc(draft.parentName)})</td></tr>
      <tr><td class="l">To:</td><td>${esc(draft.to)}</td></tr>
      <tr><td class="l">Subject:</td><td>${esc(draft.subject)}</td></tr>
      <tr><td class="l">Total:</td><td><b>${esc(draft.total)}</b></td></tr>
      ${warningRow}
    </table>
    <div class="btns">
      <a class="view" href="${esc(draft.pdfUrl)}" target="_blank">👁 View PDF</a>
      <button class="send" id="sendBtn" onclick="doSend()">📧 Send</button>
      <button onclick="doCancel()">Cancel</button>
    </div>
    <div id="status"></div>
    <script>
      const draft = ${JSON.stringify(draft)};
      function lock(msg) {
        document.getElementById('sendBtn').disabled = true;
        document.getElementById('status').textContent = msg;
      }
      function doSend() {
        lock('Sending...');
        google.script.run
          .withSuccessHandler(m => { document.getElementById('status').textContent = '✅ ' + m;
            setTimeout(() => google.script.host.close(), 1200); })
          .withFailureHandler(e => { document.getElementById('status').textContent = '❌ ' + e.message;
            document.getElementById('sendBtn').disabled = false; })
          .sendInvoiceDraft(draft);
      }
      function doCancel() {
        lock('Canceling...');
        google.script.run
          .withSuccessHandler(() => google.script.host.close())
          .withFailureHandler(() => google.script.host.close())
          .discardInvoiceDraft(draft.pdfFileId);
      }
    </script>
  `).setWidth(480).setHeight(300);

  ui.showModalDialog(html, 'Review Invoice');
}

/**
 * Sends every invoice whose "Email Sent?" column is not ✓.
 */
function sendAllUnsentInvoices() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.INVOICES);
  if (!sheet) {
    ui.alert('Sheet "' + CONFIG.SHEETS.INVOICES + '" was not found. Check CONFIG.SHEETS.');
    return;
  }

  const lastRow = sheet.getLastRow();
  let sentCount = 0;
  const errors = [];

  for (let row = CONFIG.DATA_START_ROW; row <= lastRow; row++) {
    const childName = sheet.getRange(row, CONFIG.COLUMNS.CHILD_NAME).getValue();
    const sentStatus = sheet.getRange(row, CONFIG.COLUMNS.EMAIL_SENT).getValue();

    // Skip empty rows or already-sent invoices
    if (!childName) continue;
    if (sentStatus === CONFIG.MARKS.YES || sentStatus === true) continue;

    try {
      processInvoiceRow(sheet, row); // marks "Email Sent?" itself
      sentCount++;
    } catch (error) {
      errors.push('Row ' + row + ': ' + error.message);
    }
  }

  let message = '✅ Sent ' + sentCount + ' invoice(s).';
  if (errors.length > 0) {
    message += '\n\n❌ Errors:\n' + errors.join('\n');
  }
  ui.alert(message);
}

// ==========================================
// CORE PROCESSING
// ==========================================

/**
 * Builds and emails the invoice for a single row of the 💰 Invoices sheet,
 * then saves the PDF to Drive.
 */
/**
 * Builds everything needed to send an invoice for a row, WITHOUT sending:
 * generates the PDF, saves it to the invoices folder, and returns a
 * draft object describing the email. Used by both the preview dialog
 * and the direct-send path.
 */
function prepareInvoiceDraft(sheet, row) {
  const C = CONFIG.COLUMNS;

  const childName = sheet.getRange(row, C.CHILD_NAME).getValue();
  const units = Number(sheet.getRange(row, C.UNITS).getValue()) || 0;
  const hours = sheet.getRange(row, C.HOURS).getValue();
  const total = Number(sheet.getRange(row, C.TOTAL).getValue()) || 0;

  if (!childName) throw new Error("Child's name is required (column C).");
  if (!total) throw new Error('Total Amount is required (column I).');

  const periodStartRaw = sheet.getRange(row, C.PERIOD_START).getValue();
  const periodEndRaw = sheet.getRange(row, C.PERIOD_END).getValue();
  const amountPaid = Number(sheet.getRange(row, C.AMOUNT_PAID).getValue()) || 0;

  const rowData = {
    childName: childName,
    periodStart: formatDate(periodStartRaw),
    periodEnd: formatDate(periodEndRaw),
    periodStartDate: periodStartRaw instanceof Date ? periodStartRaw : null,
    periodEndDate: periodEndRaw instanceof Date ? periodEndRaw : null,
    units: units,
    hours: hours,
    total: total,
    amountPaid: amountPaid,
    // Per-unit rate (workbook is billed in units, e.g. $111/unit)
    rate: units > 0 ? total / units : CONFIG.DEFAULT_RATE_PER_UNIT,
    // Per-hour rate - the invoice template's line item is Hours x Rate,
    // so the rate shown there must be hourly for the math to add up
    ratePerHour: Number(hours) > 0 ? total / Number(hours) : 0
  };

  // Look up the client on the 👥 Clients sheet (matched on Child's Name)
  const clientInfo = getClientInfo(childName);
  if (!clientInfo) {
    throw new Error('Client "' + childName + '" not found on the "' +
      CONFIG.SHEETS.CLIENTS + '" sheet.');
  }
  if (!clientInfo.email) {
    throw new Error('No email address on file for "' + childName + '".');
  }

  const invoiceData = {
    ...rowData,
    parentName: clientInfo.parentName,
    clientDob: clientInfo.dob,
    clientDiagnosis: clientInfo.diagnosis,
    clientEmail: clientInfo.email,
    clientAddress: clientInfo.street,
    clientCityState: clientInfo.city + ', ' + clientInfo.state + ' ' + clientInfo.zip,
    clientPhone: clientInfo.phone
  };

  // Use the Invoice # already on the row if present, otherwise generate the next one
  const invoiceNumber = getOrAssignInvoiceNumber(sheet, row);

  // Pull sessions up front so fillSessionTable doesn't fetch them twice.
  // The Invoices row total is always the authoritative billing amount -
  // if the session table lines don't add up to match it (e.g. Session Data
  // is incomplete), warn in the preview but still use the row total.
  let warning = null;
  invoiceData.sessions = getSessionsForInvoice(
    invoiceData.childName, invoiceData.periodStartDate, invoiceData.periodEndDate);
  if (invoiceData.sessions.length > 0) {
    const sessionsTotal = invoiceData.sessions.reduce((sum, s) => sum + (s.total || 0), 0);
    if (Math.abs(sessionsTotal - invoiceData.total) > 0.005) {
      warning = 'The ' + invoiceData.sessions.length + ' session(s) in Session Data for ' +
        'this period add up to ' + formatCurrency(sessionsTotal) +
        ', but the invoice total is ' + formatCurrency(invoiceData.total) +
        '. The invoice total is correct — check the Session Data sheet for missing entries.';
    }
  }

  const pdfBlob = createInvoicePdf(invoiceData, invoiceNumber);
  const email = getEmailContent(invoiceData, invoiceNumber);

  // Save the PDF now so it can be previewed; if the user cancels,
  // discardInvoiceDraft() trashes it again.
  const pdfFile = savePdfToDrive(pdfBlob, invoiceNumber, invoiceData.childName);

  return {
    sheetName: sheet.getName(),
    row: row,
    invoiceNumber: invoiceNumber,
    childName: invoiceData.childName,
    parentName: invoiceData.parentName,
    to: invoiceData.clientEmail,
    subject: email.subject,
    body: email.body,
    htmlBody: email.htmlBody,
    total: formatCurrency(invoiceData.total),
    warning: warning,
    pdfFileId: pdfFile.getId(),
    pdfUrl: pdfFile.getUrl()
  };
}

/**
 * Sends a previously prepared draft. Called by the preview dialog's
 * Send button (via google.script.run) and by the direct-send path.
 */
function sendInvoiceDraft(draft) {
  const pdfBlob = DriveApp.getFileById(draft.pdfFileId).getBlob();
  sendEmail(draft.to, draft.subject, draft.body, pdfBlob, draft.htmlBody);

  // Look the sheet up by CONFIG name rather than trusting draft.sheetName:
  // the emoji in the tab name can get mangled on the round-trip through
  // the preview dialog, which made this step silently do nothing.
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEETS.INVOICES);
  if (sheet) {
    sheet.getRange(draft.row, CONFIG.COLUMNS.EMAIL_SENT)
      .setValue(CONFIG.MARKS.YES)
      .setFontColor('#188038')
      .setHorizontalAlignment('center');
    SpreadsheetApp.flush();
  }
  return 'Invoice #' + draft.invoiceNumber + ' sent to ' + draft.to;
}

/**
 * Discards a prepared draft (preview dialog's Cancel button):
 * trashes the generated PDF.
 */
function discardInvoiceDraft(pdfFileId) {
  try {
    DriveApp.getFileById(pdfFileId).setTrashed(true);
  } catch (err) {
    // already gone - nothing to do
  }
  return 'Invoice canceled.';
}

/**
 * Prepares and immediately sends the invoice for a row (no preview).
 * Used by the column-N checkbox trigger and "Send All Unsent Invoices".
 */
function processInvoiceRow(sheet, row) {
  const draft = prepareInvoiceDraft(sheet, row);
  sendInvoiceDraft(draft);
}

/**
 * Looks up a client by Child's Name on the 👥 Clients sheet.
 * Returns null if not found. Matching is case/space-insensitive.
 */
function getClientInfo(childName) {
  const clientsSheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEETS.CLIENTS);
  if (!clientsSheet) {
    throw new Error('Sheet "' + CONFIG.SHEETS.CLIENTS + '" was not found. Check CONFIG.SHEETS.');
  }

  const K = CONFIG.CLIENT_COLUMNS;
  const data = clientsSheet.getDataRange().getValues();
  const target = String(childName).trim().toLowerCase();

  for (let i = 1; i < data.length; i++) { // skip header row
    const name = String(data[i][K.CHILD_NAME - 1] || '').trim().toLowerCase();
    if (name && name === target) {
      return {
        childName: data[i][K.CHILD_NAME - 1],
        parentName: data[i][K.PARENT_NAME - 1],
        dob: formatDate(data[i][K.DOB - 1]),
        diagnosis: formatCode(data[i][K.DIAGNOSIS - 1]),
        phone: data[i][K.PHONE - 1],
        street: data[i][K.STREET - 1],
        city: data[i][K.CITY - 1],
        state: data[i][K.STATE - 1],
        zip: data[i][K.ZIP - 1],
        // Column J may hold several addresses separated by commas -
        // MailApp accepts a comma-separated list, so normalize spacing.
        email: String(data[i][K.EMAIL - 1] || '')
          .split(',').map(s => s.trim()).filter(Boolean).join(',')
      };
    }
  }
  return null;
}

/**
 * Returns the row's Invoice # if present; otherwise assigns
 * (highest existing number + 1) and writes it to the row.
 */
function getOrAssignInvoiceNumber(sheet, row) {
  const C = CONFIG.COLUMNS;
  const existing = sheet.getRange(row, C.INVOICE_NUM).getValue();
  if (existing !== '' && existing !== null) {
    return CONFIG.INVOICE_PREFIX + existing;
  }

  const lastRow = sheet.getLastRow();
  let maxNum = 1500; // starting point if the column is empty
  if (lastRow >= CONFIG.DATA_START_ROW) {
    const nums = sheet.getRange(CONFIG.DATA_START_ROW, C.INVOICE_NUM,
      lastRow - CONFIG.DATA_START_ROW + 1, 1).getValues();
    for (const [n] of nums) {
      const v = parseInt(n, 10);
      if (!isNaN(v) && v > maxNum) maxNum = v;
    }
  }

  const next = maxNum + 1;
  sheet.getRange(row, C.INVOICE_NUM).setValue(next);
  return CONFIG.INVOICE_PREFIX + next;
}

// ==========================================
// PDF / EMAIL GENERATION
// ==========================================

/**
 * Creates the invoice PDF. Uses the template spreadsheet if
 * CONFIG.INVOICE_TEMPLATE_ID is set; otherwise builds a simple
 * invoice layout from scratch.
 */
function createInvoicePdf(data, invoiceNumber) {
  const invoiceDate = new Date();
  const dueDate = new Date();
  dueDate.setDate(dueDate.getDate() + CONFIG.PAYMENT_TERMS_DAYS);

  let tempFile;

  const templateFile = findTemplateFile(
    CONFIG.INVOICE_TEMPLATE_ID,
    CONFIG.INVOICE_TEMPLATE_NAME,
    MimeType.GOOGLE_SHEETS
  );

  if (templateFile) {
    // --- Template path: copy the template and fill in placeholders ---
    tempFile = templateFile.makeCopy('Temp Invoice ' + invoiceNumber);
    const tempSs = SpreadsheetApp.open(tempFile);
    // Use the "Invoice" tab if there is one; otherwise the first tab
    const sheet = tempSs.getSheetByName('Invoice') || tempSs.getSheets()[0];

    const balance = data.total - (data.amountPaid || 0);
    const replacements = {
      '{{BUSINESS_NAME}}': CONFIG.BUSINESS.name,
      '{{BUSINESS_PRACTITIONER}}': CONFIG.BUSINESS.practitioner,
      '{{BUSINESS_ADDRESS}}': CONFIG.BUSINESS.address,
      '{{BUSINESS_CITY_STATE}}': CONFIG.BUSINESS.cityState,
      '{{BUSINESS_LICENSE}}': CONFIG.BUSINESS.license,
      '{{BUSINESS_NPI}}': CONFIG.BUSINESS.npi,
      '{{BUSINESS_TIN}}': CONFIG.BUSINESS.tin,
      '{{BUSINESS_PHONE}}': CONFIG.BUSINESS.phone,
      '{{BUSINESS_EMAIL}}': getBusinessEmail(),
      '{{INVOICE_NUMBER}}': invoiceNumber,
      '{{INVOICE_DATE}}': formatDate(invoiceDate),
      '{{DATE}}': formatDate(invoiceDate),
      '{{DUE_DATE}}': formatDate(dueDate),
      '{{PERIOD_START}}': data.periodStart,
      '{{PERIOD_END}}': data.periodEnd,
      '{{START_DATE}}': data.periodStart,
      '{{END_DATE}}': data.periodEnd,
      '{{SERVICE_PERIOD}}': data.periodStart + ' to ' + data.periodEnd,
      '{{CLIENT_NAME}}': data.parentName || data.childName,
      '{{CHILD_NAME}}': data.childName,
      '{{PATIENT_NAME}}': data.childName,
      '{{DOB}}': data.clientDob,
      '{{DIAGNOSIS}}': data.clientDiagnosis,
      '{{CLIENT_ADDRESS}}': data.clientAddress,
      '{{ADDRESS}}': data.clientAddress,
      '{{CLIENT_CITY_STATE}}': data.clientCityState,
      '{{CITY_STATE_ZIP}}': data.clientCityState,
      '{{CLIENT_EMAIL}}': data.clientEmail,
      '{{CLIENT_PHONE}}': data.clientPhone,
      '{{UNITS}}': data.units,
      '{{HOURS}}': data.hours,
      '{{TOTAL_HOURS}}': data.hours,
      // Template line item is Hours x Rate = Amount, so RATE is hourly
      '{{RATE}}': formatCurrency(data.ratePerHour || data.rate),
      '{{TOTAL}}': formatCurrency(data.total),
      '{{TOTAL_AMOUNT}}': formatCurrency(data.total),
      '{{SUBTOTAL}}': formatCurrency(data.total),
      '{{TOTAL_DUE}}': formatCurrency(data.total),
      '{{AMOUNT_PAID}}': formatCurrency(data.amountPaid || 0),
      '{{BALANCE}}': formatCurrency(balance)
    };
    replacePlaceholders(sheet, replacements);
    fillSessionTable(sheet, data);
    SpreadsheetApp.flush();
    verifyNoPlaceholdersLeft(sheet);
  } else {
    // --- Fallback path: build a basic invoice in a temp spreadsheet ---
    const tempSs = SpreadsheetApp.create('Temp Invoice ' + invoiceNumber);
    tempFile = DriveApp.getFileById(tempSs.getId());
    const sheet = tempSs.getSheets()[0];

    const contactLine = [CONFIG.BUSINESS.phone, getBusinessEmail()]
      .filter(Boolean).join('  |  ');
    const rows = [
      [CONFIG.BUSINESS.practitioner, '', '', ''],
      [CONFIG.BUSINESS.name, '', '', ''],
      [CONFIG.BUSINESS.address, '', '', ''],
      [CONFIG.BUSINESS.cityState, '', '', ''],
      ['NY License #: ' + CONFIG.BUSINESS.license, '', '', ''],
      ['NPI #: ' + CONFIG.BUSINESS.npi, '', '', ''],
      ['TIN #: ' + CONFIG.BUSINESS.tin, '', '', ''],
      [contactLine, '', '', ''],
      ['', '', '', ''],
      ['INVOICE', '', 'Invoice #: ' + invoiceNumber, ''],
      ['', '', 'Date: ' + formatDate(invoiceDate), ''],
      ['', '', 'Due: ' + formatDate(dueDate), ''],
      ['', '', '', ''],
      ['Bill To:', '', '', ''],
      [data.parentName || data.childName, '', '', ''],
      [data.clientAddress || '', '', '', ''],
      [data.clientCityState || '', '', '', ''],
      ['', '', '', ''],
      ['Description', 'Units', 'Rate / Unit', 'Amount'],
      ['Occupational therapy services for ' + data.childName +
        ' (' + data.periodStart + ' - ' + data.periodEnd + ')',
        data.units, formatCurrency(data.rate), formatCurrency(data.total)],
      ['', '', '', ''],
      ['', '', 'TOTAL DUE:', formatCurrency(data.total)],
      ['', '', '', ''],
      ['Payment is due within ' + CONFIG.PAYMENT_TERMS_DAYS + ' days. Thank you!', '', '', '']
    ];
    sheet.getRange(1, 1, rows.length, 4).setValues(rows);

    // Format by content position so layout changes can't break it
    const invoiceRow = rows.findIndex(r => r[0] === 'INVOICE') + 1;
    const headerRow = rows.findIndex(r => r[0] === 'Description') + 1;
    const totalRow = rows.findIndex(r => r[2] === 'TOTAL DUE:') + 1;
    sheet.getRange('A1:A2').setFontWeight('bold');
    sheet.getRange(invoiceRow, 1).setFontSize(16).setFontWeight('bold');
    sheet.getRange(headerRow, 1, 1, 4).setFontWeight('bold').setBackground('#f0f4f8');
    sheet.getRange(totalRow, 3, 1, 2).setFontWeight('bold');
    sheet.setColumnWidth(1, 320);
    SpreadsheetApp.flush();
  }

  // Export the temp spreadsheet as PDF, then trash it.
  // NOTE: file.getAs(PDF) can return a STALE cached export that predates
  // the placeholder replacements (PDF with no data). The export URL with
  // an OAuth token always reflects the current content.
  const exportUrl = 'https://docs.google.com/spreadsheets/d/' + tempFile.getId() +
    '/export?format=pdf&size=letter&portrait=true&fitw=true' +
    '&gridlines=false&printtitle=false&sheetnames=false&pagenum=UNDEFINED';
  const response = UrlFetchApp.fetch(exportUrl, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
  });
  const pdfBlob = response.getBlob().setName('Invoice_' + invoiceNumber + '.pdf');
  tempFile.setTrashed(true);
  return pdfBlob;
}

/**
 * Builds the email subject and body - from the Google Doc template if
 * found, otherwise from built-in text.
 *
 * When the Doc template is used, it is exported as HTML so the email
 * keeps the template's formatting (bold, spacing); a plain-text version
 * is included as a fallback for text-only mail clients.
 *
 * If the template's first line is "EMAIL SUBJECT: ...", that line is
 * used as the email subject (placeholders filled in) and removed from
 * the body. Otherwise a default subject is used.
 *
 * Returns { subject, body, htmlBody } (htmlBody is null without a template).
 */
function getEmailContent(data, invoiceNumber) {
  const dueDate = new Date();
  dueDate.setDate(dueDate.getDate() + CONFIG.PAYMENT_TERMS_DAYS);

  const fillPlaceholders = text => text
    .replace(/\{\{INVOICE_NUMBER\}\}/g, invoiceNumber)
    .replace(/\{\{CLIENT_NAME\}\}/g, data.parentName || data.childName)
    .replace(/\{\{CHILD_NAME\}\}/g, data.childName)
    .replace(/\{\{PERIOD_START\}\}/g, data.periodStart)
    .replace(/\{\{PERIOD_END\}\}/g, data.periodEnd)
    .replace(/\{\{UNITS\}\}/g, data.units)
    .replace(/\{\{HOURS\}\}/g, data.hours)
    .replace(/\{\{TOTAL_HOURS\}\}/g, data.hours)
    // Email template's line item is Hours x Rate, so RATE is hourly -
    // same convention as the invoice PDF
    .replace(/\{\{RATE\}\}/g, formatCurrency(data.ratePerHour || data.rate))
    .replace(/\{\{TOTAL_DUE\}\}/g, formatCurrency(data.total))
    .replace(/\{\{DUE_DATE\}\}/g, formatDate(dueDate))
    .replace(/\{\{BUSINESS_NAME\}\}/g, CONFIG.BUSINESS.name)
    .replace(/\{\{BUSINESS_PRACTITIONER\}\}/g, CONFIG.BUSINESS.practitioner)
    .replace(/\{\{BUSINESS_PHONE\}\}/g, CONFIG.BUSINESS.phone)
    .replace(/\{\{BUSINESS_EMAIL\}\}/g, getBusinessEmail());

  let raw;
  let htmlBody = null;
  const templateFile = findTemplateFile(
    CONFIG.EMAIL_TEMPLATE_ID,
    CONFIG.EMAIL_TEMPLATE_NAME,
    MimeType.GOOGLE_DOCS
  );

  if (templateFile) {
    raw = DocumentApp.openById(templateFile.getId()).getBody().getText();

    // Export the Doc as HTML to preserve the template's formatting
    try {
      const htmlUrl = 'https://docs.google.com/feeds/download/documents/export/Export?id=' +
        templateFile.getId() + '&exportFormat=html';
      let html = UrlFetchApp.fetch(htmlUrl, {
        headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
      }).getContentText();
      // Remove the EMAIL SUBJECT paragraph from the HTML body
      html = html.replace(/<p[^>]*>(?:(?!<\/p>)[\s\S])*EMAIL SUBJECT:(?:(?!<\/p>)[\s\S])*<\/p>/i, '');
      htmlBody = fillPlaceholders(html);
    } catch (err) {
      htmlBody = null; // fall back to plain text
      console.error('HTML export of email template failed: ' + err.message);
    }
  } else {
    raw =
      'Dear {{CLIENT_NAME}},\n\n' +
      'Please find attached invoice #{{INVOICE_NUMBER}} for occupational therapy ' +
      'services provided for {{CHILD_NAME}} from {{PERIOD_START}} to {{PERIOD_END}}.\n\n' +
      'Total due: {{TOTAL_DUE}}\n' +
      'Due date: {{DUE_DATE}}\n\n' +
      'Please don\'t hesitate to reach out with any questions.\n\n' +
      'Thank you,\n{{BUSINESS_NAME}}\n{{BUSINESS_PHONE}}\n{{BUSINESS_EMAIL}}';
  }

  const filled = fillPlaceholders(raw);

  // Pull out an "EMAIL SUBJECT:" first line if the template has one
  let subject = 'Invoice #' + invoiceNumber + ' from ' + CONFIG.BUSINESS.name;
  let body = filled;
  const match = filled.match(/^\s*EMAIL SUBJECT:\s*(.+)\s*$/im);
  if (match) {
    subject = match[1].trim();
    body = filled.replace(match[0], '').replace(/^\s+/, '');
  }

  return { subject: subject, body: body, htmlBody: htmlBody };
}

/**
 * Sends the email with the PDF attached.
 * `toEmail` may be a comma-separated list.
 * - htmlBody (if provided) carries the template's formatting; the plain
 *   body is the fallback for text-only clients.
 * - A copy is BCC'd to the account sending it, so every invoice also
 *   lands in that inbox (not just the Sent folder).
 */
function sendEmail(toEmail, subject, body, pdfBlob, htmlBody) {
  const options = {
    to: toEmail,
    subject: subject,
    body: body,
    attachments: [pdfBlob]
  };
  if (htmlBody) options.htmlBody = htmlBody;

  const selfCopy = Session.getActiveUser().getEmail();
  if (selfCopy && String(toEmail).toLowerCase().indexOf(selfCopy.toLowerCase()) === -1) {
    options.bcc = selfCopy;
  }

  MailApp.sendEmail(options);
}

/**
 * Saves a copy of every generated invoice PDF.
 *
 * The folder is resolved automatically - no ID required:
 *   1. If CONFIG.INVOICES_FOLDER_ID is set, that exact folder is used.
 *   2. Otherwise: an "invoices" subfolder inside the SAME Drive folder
 *      as this spreadsheet (created on first use if it doesn't exist).
 *      If the spreadsheet sits in My Drive root, the subfolder is
 *      created in the root.
 */
function savePdfToDrive(pdfBlob, invoiceNumber, childName) {
  const folder = getInvoicesFolder();
  const safeName = String(childName).replace(/[^a-zA-Z0-9]/g, '_');
  return folder.createFile(pdfBlob.copyBlob().setName(invoiceNumber + '_' + safeName + '.pdf'));
}

/**
 * Returns the folder where invoice PDFs are stored, creating the
 * "invoices" subfolder next to this spreadsheet if needed.
 */
function getInvoicesFolder() {
  if (CONFIG.INVOICES_FOLDER_ID) {
    return DriveApp.getFolderById(CONFIG.INVOICES_FOLDER_ID);
  }

  // Locate the Drive folder containing this spreadsheet
  const ssFile = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  const parents = ssFile.getParents();
  const parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();

  // Reuse the "invoices" subfolder if it already exists, else create it
  const existing = parent.getFoldersByName(CONFIG.INVOICES_FOLDER_NAME);
  return existing.hasNext() ? existing.next() : parent.createFolder(CONFIG.INVOICES_FOLDER_NAME);
}

// ==========================================
// UTILITY FUNCTIONS
// ==========================================

/**
 * Returns the business email: CONFIG value if set, otherwise the
 * Google account running the script (which is also the address the
 * invoice emails are sent from).
 */
function getBusinessEmail() {
  return CONFIG.BUSINESS.email || Session.getActiveUser().getEmail() || '';
}

/**
 * Finds a template file. Order of preference:
 *   1. Exact file ID, if one is configured.
 *   2. Search Drive (excluding trash) for a Google-format file whose name
 *      matches with OR without the original extension - Drive conversions
 *      often keep the extension in the title (e.g. "invoice_template.xlsx"
 *      as a Google Sheet). Filtering by Google MIME type means an
 *      unconverted Office upload with the same name is never picked up.
 * Returns the Drive File, or null if nothing is found.
 */
function findTemplateFile(fileId, fileName, mimeType) {
  if (fileId) {
    try {
      return DriveApp.getFileById(fileId);
    } catch (err) {
      throw new Error('Template ID "' + fileId + '" is configured but the file ' +
        'could not be opened. Fix or blank out the ID in CONFIG.');
    }
  }
  if (!fileName) return null;

  const ext = mimeType === MimeType.GOOGLE_SHEETS ? '.xlsx' :
              mimeType === MimeType.GOOGLE_DOCS ? '.docx' : '';
  const candidates = ext ? [fileName, fileName + ext] : [fileName];

  const titleClause = candidates
    .map(n => 'title = "' + n.replace(/"/g, '\\"') + '"')
    .join(' or ');
  const files = DriveApp.searchFiles(
    '(' + titleClause + ') and mimeType = "' + mimeType + '" and trashed = false'
  );
  return files.hasNext() ? files.next() : null;
}

/**
 * Formats a date as MM/DD/YYYY. Returns strings unchanged.
 */
function formatDate(date) {
  if (!date) return '';
  if (typeof date === 'string') return date;
  const d = new Date(date);
  if (isNaN(d.getTime())) return String(date);
  return String(d.getMonth() + 1).padStart(2, '0') + '/' +
         String(d.getDate()).padStart(2, '0') + '/' +
         d.getFullYear();
}

/**
 * Formats a number as currency, e.g. 1554 -> $1,554.00
 */
function formatCurrency(amount) {
  const n = Number(amount);
  if (isNaN(n)) return '$0.00';
  return '$' + n.toFixed(2).replace(/\d(?=(\d{3})+\.)/g, '$&,');
}

/**
 * Replaces {{PLACEHOLDER}} text in a sheet using TextFinder, which is
 * applied server-side (no read-modify-write race with the PDF export)
 * and handles cells with mixed/rich formatting.
 */
function replacePlaceholders(sheet, replacements) {
  for (const key in replacements) {
    const value = replacements[key];
    sheet.createTextFinder(key)
      .matchCase(true)
      .replaceAllWith(value == null ? '' : String(value));
  }
}

/**
 * Throws if ANY {{...}} placeholder is still present in the sheet -
 * including ones the script has never heard of - so a half-filled
 * invoice can never be sent silently, and the error names exactly
 * which placeholders the script needs to learn.
 */
function verifyNoPlaceholdersLeft(sheet) {
  const values = sheet.getDataRange().getValues();
  const leftover = new Set();
  for (const row of values) {
    for (const cell of row) {
      if (typeof cell !== 'string') continue;
      const matches = cell.match(/\{\{[A-Za-z0-9_]+\}\}/g);
      if (matches) matches.forEach(m => leftover.add(m));
    }
  }
  if (leftover.size > 0) {
    throw new Error('The invoice template contains placeholders the script ' +
      'does not recognize: ' + Array.from(leftover).join(', ') +
      '. Add them to the replacements map in createInvoicePdf, or remove ' +
      'them from the template.');
  }
}

/**
 * Fills the "Services Delivered" table in the invoice with the client's
 * sessions for the billing period. Does nothing if the template has no
 * such table (detected by a "CPT Code" header cell).
 *
 * Sessions come from the Session Data sheet, falling back to the
 * per-child 👶 tab. Location and Mod columns use CONFIG defaults since
 * the workbook doesn't track them per-session.
 */
function fillSessionTable(sheet, data) {
  const headerCell = sheet.createTextFinder('CPT Code').matchCase(false).findNext();
  if (!headerCell) return;

  const sessions = data.sessions || getSessionsForInvoice(
    data.childName, data.periodStartDate, data.periodEndDate);
  if (sessions.length === 0) return;

  const headerRow = headerCell.getRow();
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0]
    .map(h => String(h).trim().toLowerCase());
  const col = label => headers.indexOf(label) + 1; // 0 = not found
  const cols = {
    date: col('date'),
    location: col('location'),
    cpt: headers.findIndex(h => h.indexOf('cpt') === 0) + 1,
    mod: col('mod'),
    units: col('units'),
    price: col('unit price') || col('rate'),
    total: col('total') || col('amount')
  };

  // Count the blank body rows directly under the header
  let blankRows = 0;
  const maxScan = Math.min(sheet.getLastRow() - headerRow, 100);
  const body = maxScan > 0 ?
    sheet.getRange(headerRow + 1, 1, maxScan, lastCol).getValues() : [];
  for (const r of body) {
    if (r.every(v => v === '' || v === null)) blankRows++;
    else break;
  }

  // Grow the table if there are more sessions than blank rows
  // (inserted rows inherit the formatting of the row above them)
  if (sessions.length > blankRows) {
    sheet.insertRowsAfter(headerRow + Math.max(blankRows, 1),
      sessions.length - blankRows);
  }

  sessions.forEach((s, i) => {
    const r = headerRow + 1 + i;
    if (cols.date) sheet.getRange(r, cols.date).setValue(formatDate(s.date));
    if (cols.location) sheet.getRange(r, cols.location).setValue(CONFIG.SERVICE_LOCATION);
    if (cols.cpt) sheet.getRange(r, cols.cpt).setValue(s.cpt);
    if (cols.mod) sheet.getRange(r, cols.mod).setValue(CONFIG.CPT_MODIFIER);
    if (cols.units) sheet.getRange(r, cols.units).setValue(s.units);
    if (cols.price) sheet.getRange(r, cols.price).setValue(s.rate);
    if (cols.total) sheet.getRange(r, cols.total).setValue(s.total);
  });

  // Format the filled rows: dollar amounts and center everything
  const usedCols = Object.values(cols).filter(c => c > 0);
  const firstCol = Math.min(...usedCols);
  const colSpan = Math.max(...usedCols) - firstCol + 1;
  sheet.getRange(headerRow + 1, firstCol, sessions.length, colSpan)
    .setHorizontalAlignment('center');
  if (cols.price) sheet.getRange(headerRow + 1, cols.price, sessions.length, 1)
    .setNumberFormat('$#,##0.00');
  if (cols.total) sheet.getRange(headerRow + 1, cols.total, sessions.length, 1)
    .setNumberFormat('$#,##0.00');
}

/**
 * Returns the client's sessions within the billing period (inclusive),
 * sorted by date. Reads the Session Data sheet first; if that has no
 * rows for the child, tries the per-child '👶 <name>' tab (same layout).
 */
function getSessionsForInvoice(childName, startDate, endDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const target = String(childName).trim().toLowerCase();
  const start = startDate ? new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate()) : null;
  const end = endDate ? new Date(endDate.getFullYear(), endDate.getMonth(), endDate.getDate(), 23, 59, 59) : null;

  const readSheet = name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return [];
    const values = sh.getDataRange().getValues();
    const out = [];
    for (let i = 1; i < values.length; i++) {
      const [name_, date, cpt1, cpt2, units] = values[i];
      const rate = Number(values[i][5]) || 0;
      const total = Number(values[i][6]) || (Number(units) || 0) * rate;
      if (String(name_).trim().toLowerCase() !== target) continue;
      if (!(date instanceof Date)) continue;
      if (start && date < start) continue;
      if (end && date > end) continue;
      out.push({
        date: date,
        cpt: [cpt1, cpt2].filter(v => v !== '' && v != null).map(formatCode).join(', '),
        units: Number(units) || 0,
        rate: rate,
        total: total
      });
    }
    return out;
  };

  let sessions = readSheet(CONFIG.SHEETS.SESSION_DATA);
  if (sessions.length === 0) sessions = readSheet('👶 ' + childName);
  sessions.sort((a, b) => a.date - b.date);
  return sessions;
}

/**
 * Formats numeric codes (CPT, diagnosis) without decimals:
 * 97112.0 -> "97112". Strings pass through unchanged.
 */
function formatCode(value) {
  if (value === '' || value == null) return '';
  const n = Number(value);
  return isNaN(n) ? String(value) : String(Math.round(n));
}
