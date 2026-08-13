/**
 * Weekly Sales Report — scheduled rebuild and email.
 *
 * Paste this into Extensions ▸ Apps Script in your own Google Sheet. It lives
 * in your Google account, runs on your schedule, and sends from your address.
 * Nothing here calls an outside service, so there is no subscription and
 * nothing that stops working if I disappear.
 *
 * Setup, once:
 *   1. Extensions ▸ Apps Script, paste this file over Code.gs, Save.
 *   2. Edit CONFIG below — at minimum, RECIPIENTS.
 *   3. Run `installWeeklyTrigger` once and approve the permission prompt.
 *      Google asks for permission to read this Sheet and send mail as you.
 *   4. Run `sendWeeklyReportNow` to check the email looks right.
 *
 * To change the schedule or recipients later, edit CONFIG and run
 * `installWeeklyTrigger` again. It removes its own old trigger first, so
 * running it twice never gives you two emails.
 */

const CONFIG = {
  // Who receives it. Comma-separated for several people.
  RECIPIENTS: 'owner@example.com',

  // Monday 8am, in the timezone set under File ▸ Settings for this Sheet.
  SEND_DAY: ScriptApp.WeekDay.MONDAY,
  SEND_HOUR: 8,

  // Tab names. Change these only if you rename the tabs.
  REPORT_SHEET: 'Weekly Report',
  CALC_SHEET: 'Weekly Calc',
  RAW_SHEET: 'Raw Orders',

  // Where the report keeps its reporting-week date.
  REPORTING_WEEK_CELL: 'H3',

  // Subject line. {week} is replaced with the week starting date.
  SUBJECT: 'Weekly sales report — week starting {week}',

  // Refuse to send if the newest order is older than this many days. Stops the
  // script cheerfully emailing last month's numbers when an export fails.
  STALE_AFTER_DAYS: 10,
};

/** Trigger entry point. Keep the name stable; the trigger refers to it. */
function sendWeeklyReport() {
  const result = buildAndSend_();
  console.log(result);
}

/** Same thing, for running by hand while you are setting it up. */
function sendWeeklyReportNow() {
  const result = buildAndSend_();
  SpreadsheetApp.getActiveSpreadsheet().toast(result, 'Weekly report', 8);
  console.log(result);
}

/**
 * Install (or reinstall) the weekly trigger.
 * Removes any trigger this script previously created, so running it more than
 * once cannot produce duplicate emails.
 */
function installWeeklyTrigger() {
  removeWeeklyTrigger();
  ScriptApp.newTrigger('sendWeeklyReport')
    .timeBased()
    .onWeekDay(CONFIG.SEND_DAY)
    .atHour(CONFIG.SEND_HOUR)
    .create();
  const msg = 'Weekly trigger installed: ' + dayName_(CONFIG.SEND_DAY) +
              ' at ' + CONFIG.SEND_HOUR + ':00.';
  SpreadsheetApp.getActiveSpreadsheet().toast(msg, 'Scheduled', 8);
  console.log(msg);
}

/** Turn the schedule off without deleting the script. */
function removeWeeklyTrigger() {
  let removed = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sendWeeklyReport') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  console.log('Removed ' + removed + ' existing trigger(s).');
  return removed;
}

// ───────────────────────────────────────────────────────────────────────────

function buildAndSend_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();

  const report = mustGetSheet_(ss, CONFIG.REPORT_SHEET);
  const calc = mustGetSheet_(ss, CONFIG.CALC_SHEET);
  const raw = mustGetSheet_(ss, CONFIG.RAW_SHEET);

  // Formulas are recalculated lazily. Without this the script can read the
  // values from before the newest paste.
  SpreadsheetApp.flush();

  const week = calc.getRange(CONFIG.REPORTING_WEEK_CELL).getValue();
  if (!(week instanceof Date)) {
    throw new Error('No reporting week found in ' + CONFIG.CALC_SHEET + '!' +
                    CONFIG.REPORTING_WEEK_CELL + '. Is the Raw Orders tab empty?');
  }

  // A stale check, because a scheduled job that quietly emails old numbers is
  // worse than one that fails. Better a loud error than a confident wrong figure.
  const ageDays = Math.floor((Date.now() - week.getTime()) / 86400000);
  if (ageDays > CONFIG.STALE_AFTER_DAYS) {
    throw new Error('Latest data is ' + ageDays + ' days old (week starting ' +
                    Utilities.formatDate(week, tz, 'd MMM yyyy') +
                    '). Nothing sent — paste a fresh export into ' +
                    CONFIG.RAW_SHEET + '.');
  }

  const kpis = {
    revenue: report.getRange('B6').getValue(),
    orders: report.getRange('D6').getValue(),
    aov: report.getRange('F6').getValue(),
    units: report.getRange('H6').getValue(),
    wow: report.getRange('J6').getValue(),
  };

  const products = readBlock_(report, 30, 2, 8, 3);   // name, revenue, units
  const channels = readBlock_(report, 30, 6, 5, 3);   // name, revenue, share

  const weekLabel = Utilities.formatDate(week, tz, 'd MMM yyyy');
  const subject = CONFIG.SUBJECT.replace('{week}', weekLabel);
  const html = renderEmail_(weekLabel, kpis, products, channels, ss.getUrl());

  MailApp.sendEmail({
    to: CONFIG.RECIPIENTS,
    subject: subject,
    htmlBody: html,
    body: plainText_(weekLabel, kpis),   // for clients that block HTML
    name: 'Weekly Sales Report',
  });

  return 'Sent to ' + CONFIG.RECIPIENTS + ' for week starting ' + weekLabel + '.';
}

function mustGetSheet_(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh) {
    throw new Error('Tab "' + name + '" not found. If you renamed it, update ' +
                    'CONFIG at the top of this script.');
  }
  return sh;
}

/** Read a rectangular block, dropping blank rows. */
function readBlock_(sheet, startRow, startCol, numRows, numCols) {
  return sheet.getRange(startRow, startCol, numRows, numCols)
    .getValues()
    .filter(function (r) { return r[0] !== '' && r[0] !== null; });
}

function money_(v) {
  const n = Number(v) || 0;
  return '£' + n.toLocaleString('en-GB', { minimumFractionDigits: 0,
                                           maximumFractionDigits: 0 });
}

function pct_(v) {
  const n = (Number(v) || 0) * 100;
  return (n >= 0 ? '+' : '') + n.toFixed(1) + '%';
}

function dayName_(d) {
  for (const k in ScriptApp.WeekDay) {
    if (ScriptApp.WeekDay[k] === d) return k.charAt(0) + k.slice(1).toLowerCase();
  }
  return String(d);
}

function plainText_(weekLabel, k) {
  return [
    'Weekly sales report — week starting ' + weekLabel,
    '',
    'Revenue:         ' + money_(k.revenue) + '  (' + pct_(k.wow) + ' vs last week)',
    'Orders:          ' + (Number(k.orders) || 0),
    'Avg order value: ' + money_(k.aov),
    'Units sold:      ' + (Number(k.units) || 0),
    '',
    'Revenue counts fulfilled product lines only. Shipping, discounts,',
    'cancellations and refunds are excluded and listed separately in the Sheet.',
  ].join('\n');
}

function renderEmail_(weekLabel, k, products, channels, url) {
  const up = (Number(k.wow) || 0) >= 0;
  const trendColour = up ? '#157F51' : '#B42318';

  const kpiCell = function (label, value, note) {
    return '<td style="padding:14px 18px;background:#EAF1FE;border-radius:8px;' +
           'vertical-align:top;">' +
           '<div style="font:600 10px Arial;letter-spacing:.06em;color:#6B7A8F;' +
           'text-transform:uppercase;">' + label + '</div>' +
           '<div style="font:700 22px Arial;color:#16202C;padding-top:4px;">' + value + '</div>' +
           '<div style="font:400 11px Arial;color:#6B7A8F;padding-top:2px;">' + note + '</div>' +
           '</td><td style="width:10px;"></td>';
  };

  const rows = function (data, cols) {
    return data.map(function (r, i) {
      const bg = i % 2 ? '#F6F8FB' : '#FFFFFF';
      let cells = '<td style="padding:7px 10px;font:400 13px Arial;color:#16202C;">' +
                  r[0] + '</td>';
      cells += '<td style="padding:7px 10px;font:400 13px Arial;color:#16202C;' +
               'text-align:right;">' + money_(r[1]) + '</td>';
      if (cols > 2) {
        const third = typeof r[2] === 'number' && r[2] <= 1
          ? (r[2] * 100).toFixed(1) + '%'
          : Math.round(Number(r[2]) || 0);
        cells += '<td style="padding:7px 10px;font:400 13px Arial;color:#6B7A8F;' +
                 'text-align:right;">' + third + '</td>';
      }
      return '<tr style="background:' + bg + ';">' + cells + '</tr>';
    }).join('');
  };

  const th = function (labels) {
    return '<tr>' + labels.map(function (l, i) {
      return '<th style="padding:7px 10px;font:700 10px Arial;letter-spacing:.05em;' +
             'text-transform:uppercase;color:#FFFFFF;background:#1F6FEB;text-align:' +
             (i === 0 ? 'left' : 'right') + ';">' + l + '</th>';
    }).join('') + '</tr>';
  };

  return [
    '<div style="background:#F6F8FB;padding:24px;font-family:Arial,sans-serif;">',
    '<div style="max-width:640px;margin:0 auto;background:#FFFFFF;border:1px solid #DDE3EA;',
    'border-radius:12px;padding:26px 28px;">',

    '<div style="font:700 20px Arial;color:#16202C;">Weekly Sales Report</div>',
    '<div style="font:400 12px Arial;color:#6B7A8F;padding:4px 0 18px;',
    'border-bottom:1px solid #DDE3EA;">Week starting ' + weekLabel + '</div>',

    '<table cellpadding="0" cellspacing="0" style="margin:18px 0 6px;"><tr>',
    kpiCell('Revenue', money_(k.revenue),
            '<span style="color:' + trendColour + ';font-weight:700;">' +
            pct_(k.wow) + '</span> vs last week'),
    kpiCell('Orders', Number(k.orders) || 0, 'distinct orders'),
    kpiCell('Avg order', money_(k.aov), 'per order'),
    '</tr></table>',

    '<div style="font:700 13px Arial;color:#16202C;padding:22px 0 8px;">Top products</div>',
    '<table cellpadding="0" cellspacing="0" width="100%" style="border-collapse:collapse;">',
    th(['Product', 'Revenue', 'Units']), rows(products, 3), '</table>',

    '<div style="font:700 13px Arial;color:#16202C;padding:22px 0 8px;">By channel</div>',
    '<table cellpadding="0" cellspacing="0" width="100%" style="border-collapse:collapse;">',
    th(['Channel', 'Revenue', 'Share']), rows(channels, 3), '</table>',

    '<div style="font:400 11px Arial;color:#6B7A8F;padding:20px 0 0;',
    'border-top:1px solid #DDE3EA;margin-top:22px;">',
    'Revenue counts fulfilled product lines only. Shipping, discounts, cancellations ',
    'and refunds are excluded and listed separately in the Sheet, so this figure ',
    'reconciles to the raw export.<br><br>',
    '<a href="' + url + '" style="color:#1F6FEB;text-decoration:none;font-weight:700;">',
    'Open the full report →</a></div>',

    '</div></div>',
  ].join('');
}
