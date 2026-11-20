/**
 * PORTFOLIO SAMPLE - Order & Payment Follow-up Tracker automation (Google Apps Script)
 * Works with the "Orders" and "Dashboard" sheets in Sample_02_Order_Payment_Tracker.xlsx
 * after it is imported into Google Sheets. All data in the sample is fictional.
 *
 * What it automates:
 *  1. Stamps "Last Updated" whenever order status, payment status or paid amount changes.
 *  2. Marking a payment "Paid" fills the Paid Amount automatically.
 *  3. Highlights overdue unpaid orders in red every morning.
 *  4. Emails you a daily digest of overdue payments (customer, phone, balance, days overdue).
 *  5. Builds a month-wise summary sheet (orders, sales, collected, outstanding) in one click.
 */

const CFG = {
  sheet: 'Orders',
  headerRow: 1,
  col: {               // 1-based column numbers in the Orders sheet
    orderId: 1, orderDate: 2, customer: 3, phone: 4, item: 5, amount: 6,
    orderStatus: 7, paymentStatus: 8, paid: 9, balance: 10, dueDate: 11,
    daysOverdue: 12, lastUpdated: 13, notes: 14
  },
  overdueColour: '#F8D7DA',
  summarySheet: 'Monthly_Summary',
  digestHour: 9        // daily email around 9 AM (script time zone)
};

/** Adds the Automation menu when the spreadsheet opens. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Automation')
    .addItem('Highlight overdue payments now', 'highlightOverdue')
    .addItem('Build monthly summary', 'buildMonthlySummary')
    .addItem('Email me the overdue digest now', 'emailOverdueDigest')
    .addSeparator()
    .addItem('Turn on daily 9 AM automation', 'installDailyTrigger')
    .addItem('Turn off daily automation', 'removeDailyTrigger')
    .addToUi();
}

/** Simple trigger: timestamps edits and auto-fills paid amount. */
function onEdit(e) {
  const range = e.range;
  const sh = range.getSheet();
  if (sh.getName() !== CFG.sheet || range.getRow() <= CFG.headerRow) return;

  const c = CFG.col;
  const firstCol = range.getColumn();
  const lastCol = firstCol + range.getNumColumns() - 1;
  const watched = [c.orderStatus, c.paymentStatus, c.paid];
  if (!watched.some(w => w >= firstCol && w <= lastCol)) return;

  for (let r = range.getRow(); r < range.getRow() + range.getNumRows(); r++) {
    if (!sh.getRange(r, c.orderId).getValue()) continue;
    if (String(sh.getRange(r, c.paymentStatus).getValue()).trim().toLowerCase() === 'paid') {
      sh.getRange(r, c.paid).setValue(sh.getRange(r, c.amount).getValue());
    }
    sh.getRange(r, c.lastUpdated).setValue(new Date());
  }
}

/** Reads all order rows as objects. */
function getOrders_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.sheet);
  const last = sh.getLastRow();
  if (last <= CFG.headerRow) return { sh, rows: [] };
  const values = sh.getRange(CFG.headerRow + 1, 1, last - CFG.headerRow, CFG.col.notes).getValues();
  const c = CFG.col;
  const today = startOfDay_(new Date());
  const rows = values
    .map((v, i) => {
      const amount = Number(v[c.amount - 1]) || 0;
      const paid = Number(v[c.paid - 1]) || 0;
      const due = v[c.dueDate - 1] instanceof Date ? startOfDay_(v[c.dueDate - 1]) : null;
      const balance = amount - paid;
      const cancelled = String(v[c.orderStatus - 1]).toLowerCase() === 'cancelled';
      const overdueDays = (!cancelled && balance > 0 && due && due < today)
        ? Math.round((today - due) / 86400000) : 0;
      return {
        row: CFG.headerRow + 1 + i, id: v[c.orderId - 1], date: v[c.orderDate - 1],
        customer: v[c.customer - 1], phone: v[c.phone - 1], item: v[c.item - 1],
        amount, paid, balance, due, overdueDays, cancelled,
        status: v[c.orderStatus - 1], payStatus: v[c.paymentStatus - 1]
      };
    })
    .filter(o => o.id !== '' && o.id !== null);
  return { sh, rows };
}

/** Colours overdue unpaid rows; clears colour on the rest. */
function highlightOverdue() {
  const { sh, rows } = getOrders_();
  rows.forEach(o => {
    sh.getRange(o.row, 1, 1, CFG.col.notes)
      .setBackground(o.overdueDays > 0 ? CFG.overdueColour : null);
  });
  return rows.filter(o => o.overdueDays > 0).length;
}

/** Sends the overdue digest to the account running the script. */
function emailOverdueDigest() {
  const { rows } = getOrders_();
  const overdue = rows.filter(o => o.overdueDays > 0).sort((a, b) => b.overdueDays - a.overdueDays);
  const to = Session.getEffectiveUser().getEmail();
  const ssName = SpreadsheetApp.getActive().getName();
  if (!overdue.length) {
    MailApp.sendEmail(to, `No overdue payments - ${ssName}`, 'All payments are on time today.');
    return;
  }
  const total = overdue.reduce((s, o) => s + o.balance, 0);
  const fmt = n => '₹' + Math.round(n).toLocaleString('en-IN');
  const rowsHtml = overdue.map(o =>
    `<tr><td>${o.id}</td><td>${o.customer}</td><td>${o.phone}</td>` +
    `<td style="text-align:right">${fmt(o.balance)}</td><td style="text-align:right">${o.overdueDays}</td></tr>`
  ).join('');
  const html =
    `<p>${overdue.length} overdue payment(s), total outstanding <b>${fmt(total)}</b>.</p>` +
    `<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-family:Arial;font-size:13px">` +
    `<tr style="background:#1F3A5A;color:#fff"><th>Order</th><th>Customer</th><th>Phone</th><th>Balance</th><th>Days overdue</th></tr>` +
    rowsHtml + `</table><p style="color:#777">Sent automatically by ${ssName}.</p>`;
  MailApp.sendEmail({ to, subject: `${overdue.length} overdue payments (${fmt(total)}) - ${ssName}`, htmlBody: html });
}

/** Builds / refreshes a month-wise summary sheet. */
function buildMonthlySummary() {
  const ss = SpreadsheetApp.getActive();
  const { rows } = getOrders_();
  const tz = ss.getSpreadsheetTimeZone();
  const byMonth = {};
  rows.filter(o => !o.cancelled && o.date instanceof Date).forEach(o => {
    const key = Utilities.formatDate(o.date, tz, 'yyyy-MM');
    const m = byMonth[key] || (byMonth[key] = { orders: 0, sales: 0, collected: 0, outstanding: 0, overdue: 0 });
    m.orders++; m.sales += o.amount; m.collected += o.paid; m.outstanding += o.balance;
    if (o.overdueDays > 0) m.overdue += o.balance;
  });
  const keys = Object.keys(byMonth).sort();
  const out = [['Month', 'Orders', 'Sales', 'Collected', 'Outstanding', 'Overdue', 'Collection %']];
  keys.forEach(k => {
    const m = byMonth[k];
    out.push([k, m.orders, m.sales, m.collected, m.outstanding, m.overdue, m.sales ? m.collected / m.sales : 0]);
  });
  let sh = ss.getSheetByName(CFG.summarySheet) || ss.insertSheet(CFG.summarySheet);
  sh.clear();
  sh.getRange(1, 1, out.length, out[0].length).setValues(out);
  sh.getRange(1, 1, 1, out[0].length).setFontWeight('bold').setBackground('#1F3A5A').setFontColor('#FFFFFF');
  if (out.length > 1) {
    sh.getRange(2, 3, out.length - 1, 4).setNumberFormat('₹#,##,##0');
    sh.getRange(2, 7, out.length - 1, 1).setNumberFormat('0.0%');
  }
  sh.getRange(out.length + 2, 1).setValue('Refreshed: ' + Utilities.formatDate(new Date(), tz, 'dd-MMM-yyyy HH:mm'));
  sh.autoResizeColumns(1, out[0].length);
}

/** Runs every morning once the trigger is installed. */
function dailyJob() {
  highlightOverdue();
  emailOverdueDigest();
  buildMonthlySummary();
}

function installDailyTrigger() {
  removeDailyTrigger();
  ScriptApp.newTrigger('dailyJob').timeBased().everyDays(1).atHour(CFG.digestHour).create();
  SpreadsheetApp.getActive().toast('Daily automation is on (around 9 AM).', 'Automation', 5);
}

function removeDailyTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'dailyJob')
    .forEach(t => ScriptApp.deleteTrigger(t));
}

function startOfDay_(d) {
  const x = new Date(d);
  x.setHours(0, 0, 0, 0);
  return x;
}
