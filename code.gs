/**
 * Zuru World dashboard + Monday digest.
 * Bound to the "Zuru World Hub" Google Sheet. Reads values the sheet already
 * calculates (Company Report, Faith Summary, Shem Guides, KPIs, Requests, Faith).
 * It never writes to the sheet.
 *
 * If you move a table, change the ranges in getData() below.
 */

const CONFIG = {
  // Who gets the Monday email. Replace with real addresses.
  DIGEST_TO: ['ceo@yourcompany.com'],
  DIGEST_HOUR: 8,            // 8am in the script's time zone (set Africa/Nairobi in Project Settings)
  COMPANY_NAME: 'Zuru World',
  // The dashboard and email only report on these people.
  REPORT_PEOPLE: ['Faith', 'Shem', 'Esther']
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Dashboard')
    .setTitle(CONFIG.COMPANY_NAME + ' dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Called by the dashboard page. Returns plain, serialisable data. */
function getData() {
  const ss = SpreadsheetApp.getActive();
  const tz = Session.getScriptTimeZone();
  const str = v => (v instanceof Date) ? Utilities.formatDate(v, tz, 'd MMM yyyy') : (v === null || v === undefined ? '' : v);
  const grab = (name, a1) => ss.getSheetByName(name).getRange(a1).getValues();

  // Team table: Company Report B15:L19 -> B person, D role, F tasks, G done, H overdue, I progress, L status
  const team = grab('Company Report', 'B15:L19')
    .filter(r => CONFIG.REPORT_PEOPLE.indexOf(r[0]) !== -1)
    .map(r => ({ person: r[0], role: r[2], tasks: r[4], done: r[5], overdue: r[6], progress: r[7], status: r[10] }));

  // Outreach summary: Faith Summary B4:B10
  const s = grab('Faith Summary', 'B4:B10').map(r => r[0]);
  const outreach = { listed: s[0], contacted: s[1], notContacted: s[2], followUpSent: s[3],
                     followUpNotSent: s[4], followUpDue: s[5], replies: s[6] };
  const categories = grab('Faith Summary', 'A14:D33')
    .filter(r => r[0] !== '')
    .map(r => ({ name: r[0], listed: r[1], contacted: r[2], followUp: r[3] }));

  // Faith tracker rows: B company, C category, G status, L follow-up check, M response, N next step
  const faith = grab('Faith', 'A2:O187').filter(r => r[1] !== '');
  const dueNow = faith.filter(r => r[11] === 'Follow up now').map(r => ({ company: r[1], category: r[2] }));
  const hotLeads = faith
    .filter(r => /^Replied - interested|^Call\/meeting scheduled/.test(String(r[12])))
    .map(r => ({ company: r[1], category: r[2], response: r[12], nextStep: r[13] }));

  // Guides: Shem Guides C6:C15
  const g = grab('Shem Guides', 'C6:C15').map(r => r[0]);
  const guides = { target: g[0], targetDate: str(g[1]), onboard: g[2], stillNeeded: g[3], progress: g[4],
                   daysLeft: g[5], neededPerWeek: g[6], weeklyGoal: g[7], projected: g[8], status: g[9] };

  // KPIs: A5:I24 -> person, dept, metric, actual, target, updated, progress, bar, status
  const kpis = grab('KPIs', 'A5:I24').filter(r => r[2] !== '' && CONFIG.REPORT_PEOPLE.indexOf(r[0]) !== -1)
    .map(r => ({ person: r[0], dept: r[1], metric: r[2], actual: r[3], target: r[4],
                 updated: str(r[5]), progress: r[6], status: r[8] }));

  // Requests: A5:J104 -> B by, C to, D task, E urgency, F deadline, G status, J check
  const requests = grab('Requests', 'A5:J104')
    .filter(r => r[3] !== '' && r[6] !== 'Done' &&
      (CONFIG.REPORT_PEOPLE.indexOf(r[1]) !== -1 || CONFIG.REPORT_PEOPLE.indexOf(r[2]) !== -1))
    .map(r => ({ by: r[1], to: r[2], task: r[3], urgency: r[4], deadline: str(r[5]), check: r[9] }));

  const data = { team, outreach, categories, dueNow, hotLeads, guides, kpis, requests,
                 updatedAt: Utilities.formatDate(new Date(), tz, "EEE d MMM yyyy, HH:mm") };
  data.attention = buildAttention(data);
  return data;
}

/** Plain-language list of what needs a person's attention today. */
function buildAttention(d) {
  const items = [];
  const g = d.guides;
  if (g.status === 'Below pace' || g.status === 'Deadline passed') {
    items.push({ level: 'red', text: 'Guides: ' + g.onboard + ' of ' + g.target + ' onboard with ' + g.daysLeft +
      ' days to ' + g.targetDate + '. At the current weekly goal (' + g.weeklyGoal + ') we reach about ' +
      g.projected + '; hitting the target needs about ' + g.neededPerWeek + ' a week.' });
  }
  d.hotLeads.forEach(l => {
    if (!l.nextStep) items.push({ level: 'amber', text: l.company + ' replied interested but has no next step set.' });
  });
  if (d.outreach.followUpDue > 0) {
    items.push({ level: 'amber', text: d.outreach.followUpDue + ' client follow-ups are due now (first email 7+ days ago).' });
  }
  if (d.outreach.notContacted > 0) {
    items.push({ level: 'amber', text: d.outreach.notContacted + ' listed companies have not been contacted yet.' });
  }
  const noDeadline = d.requests.filter(r => r.check === 'Add deadline').length;
  const late = d.requests.filter(r => r.check === 'Overdue').length;
  if (late) items.push({ level: 'red', text: late + ' colleague request(s) are overdue.' });
  if (noDeadline) items.push({ level: 'amber', text: noDeadline + ' open colleague request(s) have no deadline, so they cannot show as overdue.' });
  d.team.forEach(t => {
    if (t.person !== 'Faith' && t.person !== 'Shem' && t.overdue > 0)
      items.push({ level: 'amber', text: t.person + ' has ' + t.overdue + ' overdue task(s).' });
  });
  return items;
}

/* ---------- Monday email ---------- */

function sendMondayDigest() {
  const d = getData();
  const url = ScriptApp.getService().getUrl();
  const dot = { red: '#C0392B', amber: '#C98A0B', grey: '#8A97A4' };
  const li = d.attention.map(a =>
    '<li style="margin:0 0 8px"><span style="color:' + dot[a.level] + '">&#9679;</span> ' + esc_(a.text) + '</li>').join('');
  const rows = d.team.map(t =>
    '<tr><td style="padding:4px 12px 4px 0">' + esc_(t.person) + '</td><td style="padding:4px 12px 4px 0">' +
    Math.round((t.progress || 0) * 100) + '%</td><td style="padding:4px 0">' + esc_(t.status) + '</td></tr>').join('');
  const o = d.outreach;
  const html =
    '<div style="font-family:Arial,sans-serif;color:#12263A;max-width:560px">' +
    '<h2 style="margin:0 0 4px">' + CONFIG.COMPANY_NAME + ' weekly summary</h2>' +
    '<p style="margin:0 0 16px;color:#5B6B7B">' + esc_(d.updatedAt) + '</p>' +
    '<h3 style="margin:0 0 8px">Needs attention</h3><ul style="padding-left:18px;margin:0 0 16px">' +
    (li || '<li>Nothing flagged this week.</li>') + '</ul>' +
    '<h3 style="margin:0 0 8px">Outreach</h3><p style="margin:0 0 16px">' + o.contacted + ' of ' + o.listed +
    ' companies contacted, ' + o.followUpSent + ' follow-ups sent, ' + o.replies + ' replies.</p>' +
    '<h3 style="margin:0 0 8px">Team progress</h3><table style="border-collapse:collapse;margin:0 0 16px">' + rows + '</table>' +
    (url ? '<p><a href="' + url + '">Open the live dashboard</a></p>' : '') + '</div>';
  MailApp.sendEmail({
    to: CONFIG.DIGEST_TO.join(','),
    subject: CONFIG.COMPANY_NAME + ' weekly summary: ' + d.attention.filter(a => a.level !== 'grey').length + ' items need attention',
    htmlBody: html
  });
}

/** Run once from the editor to schedule the email every Monday. */
function setUpMondayDigest() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'sendMondayDigest')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('sendMondayDigest').timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(CONFIG.DIGEST_HOUR).create();
}

function esc_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
