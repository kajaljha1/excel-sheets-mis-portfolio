/**
 * Real Estate Lead & Follow-up Tracker: Google Apps Script automation
 * PORTFOLIO SAMPLE (works with Sample_04_Real_Estate_Lead_Tracker.xlsx after importing it into Google Sheets)
 *
 * What it does
 *  1. Custom menu "Lead Tracker" in the sheet.
 *  2. Every morning (9 AM) it emails the call list: overdue + due-today follow-ups, grouped by agent.
 *  3. When a new lead's Name is typed, it fills the Lead ID and Date Received automatically.
 *  4. When a Stage is changed, it stamps Last Contact with today's date.
 *
 * Setup (one time): Extensions > Apps Script > paste this file > Save > reload the sheet >
 * Lead Tracker > "Turn on daily 9 AM email". Google asks for permission once. No passwords are shared.
 */

const CFG = {
  sheet: 'Leads',
  col: { id: 1, date: 2, name: 3, phone: 4, source: 5, req: 6, loc: 7, budget: 8, stage: 9, agent: 10,
         lastContact: 11, nextFU: 12 },
  closedStages: ['Booked', 'Lost'],
  // Optional: agent name -> email. Agents listed here also get their own list.
  agentEmails: { /* 'Rahul': 'rahul@example.com' */ },
  hour: 9
};

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Lead Tracker')
    .addItem("Email today's call list now", 'sendFollowUpDigest')
    .addItem('Turn on daily 9 AM email', 'createDailyTrigger')
    .addItem('Turn off daily email', 'removeTriggers')
    .addToUi();
}

/** Builds the overdue / due-today list and emails it. */
function sendFollowUpDigest() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.sheet);
  const data = sh.getDataRange().getValues();
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const tz = Session.getScriptTimeZone();
  const due = [];

  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const name = row[CFG.col.name - 1];
    const stage = row[CFG.col.stage - 1];
    const fu = row[CFG.col.nextFU - 1];
    if (!name || CFG.closedStages.indexOf(stage) > -1 || !(fu instanceof Date)) continue;
    const d = new Date(fu); d.setHours(0, 0, 0, 0);
    if (d <= today) {
      due.push({
        name: name, phone: row[CFG.col.phone - 1], stage: stage, req: row[CFG.col.req - 1],
        budget: row[CFG.col.budget - 1], agent: row[CFG.col.agent - 1] || 'Unassigned',
        date: Utilities.formatDate(d, tz, 'dd-MMM'), late: Math.round((today - d) / 86400000)
      });
    }
  }
  if (!due.length) return;
  due.sort((a, b) => b.late - a.late);

  const table = list => '<table style="border-collapse:collapse;font-family:Arial;font-size:13px">' +
    '<tr style="background:#1F3A5A;color:#fff"><th align="left">Name</th><th>Phone</th><th>Stage</th><th>Need</th>' +
    '<th>Budget</th><th>Follow-up</th><th>Days late</th><th>Agent</th></tr>' +
    list.map(x => `<tr style="border-bottom:1px solid #ddd"><td>${x.name}</td><td>${x.phone}</td><td>${x.stage}</td>` +
      `<td>${x.req}</td><td>₹ ${x.budget} L</td><td>${x.date}</td>` +
      `<td style="color:${x.late > 0 ? '#C62828' : '#9C5900'};font-weight:bold">${x.late}</td><td>${x.agent}</td></tr>`).join('') +
    '</table>';

  const owner = Session.getEffectiveUser().getEmail();
  const subject = `Today's call list: ${due.length} follow-ups (${due.filter(x => x.late > 0).length} overdue)`;
  MailApp.sendEmail({ to: owner, subject: subject,
    htmlBody: `<p>Good morning! Here are today's follow-ups, most overdue first:</p>${table(due)}` +
              `<p style="color:#666">Sent automatically by your Lead Tracker.</p>` });

  Object.keys(CFG.agentEmails).forEach(agent => {
    const mine = due.filter(x => x.agent === agent);
    if (mine.length) MailApp.sendEmail({ to: CFG.agentEmails[agent],
      subject: `Your call list today: ${mine.length} follow-ups`, htmlBody: table(mine) });
  });
}

/** Auto Lead ID + Date Received for new leads; Last Contact stamp on stage change. */
function onEdit(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== CFG.sheet || e.range.getRow() < 2) return;
  const r = e.range.getRow(), c = e.range.getColumn();

  if (c === CFG.col.name && e.value && !sh.getRange(r, CFG.col.id).getValue()) {
    const ids = sh.getRange(2, CFG.col.id, Math.max(sh.getLastRow() - 1, 1), 1).getValues().flat()
      .map(v => parseInt(String(v).replace('L-', ''), 10)).filter(n => !isNaN(n));
    const next = (ids.length ? Math.max.apply(null, ids) : 0) + 1;
    sh.getRange(r, CFG.col.id).setValue('L-' + ('000' + next).slice(-3));
    sh.getRange(r, CFG.col.date).setValue(new Date());
  }
  if (c === CFG.col.stage && e.value) {
    sh.getRange(r, CFG.col.lastContact).setValue(new Date());
  }
}

function createDailyTrigger() {
  removeTriggers();
  ScriptApp.newTrigger('sendFollowUpDigest').timeBased().everyDays(1).atHour(CFG.hour).create();
  SpreadsheetApp.getUi().alert('Done: the call list will be emailed every day around ' + CFG.hour + ' AM.');
}

function removeTriggers() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'sendFollowUpDigest')
    .forEach(t => ScriptApp.deleteTrigger(t));
}
