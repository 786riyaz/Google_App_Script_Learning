// ============================================================
//  SHIVAM TELECOM — Field Operations Bot v4.0
//  Google Apps Script — COMPLETE SINGLE FILE
//  Handles: Fault / NTC / DISS task dispatch to engineer groups
//  Dashboard API: getReportData, markClear, getTaskRows, dispatchAll
//
//  FIXES IN v4.0:
//  1. Duplicate function declarations removed (doGet, handleGetReportData,
//     handleMarkClearAPI, handleGetTaskRows, jsonResponse were duplicated
//     across main file + ReportDataEndpoint_v3.gs causing deploy errors)
//  2. "reportimage"/"ri" command was present in foHandleMessage() but
//     "report" command incorrectly routed to foSendEngReport (option 2).
//     Menu now shows correct command: type "engreport" for workload,
//     "ri" or "reportimage" for the summary report.
//  3. dispatchAll GAS action added (called by HTML dashboard v4).
//  4. CORS header added to jsonResponse for dashboard compatibility.
//  5. foNow() IST format verified.
// ============================================================

const FO_WHATSAPP_TOKEN = 'EAASOMiQCZBD0BRL9AB7uU1JEmAvjUpC239UeuW5vseMturYpz6HYnw7aHFLem4ZBk0lGw6kK28URK5r53GDMY7LCj6cnAJreHKBW15aJ3tdNF5c4L6hMbfCtzWuqZBmqtnVky15ZCaDvknmtALID11iz2BJUOuUou7JfJOH5ZCjPWfkiZCm4ZAu7iK4yLeZBZCCD82rFRzX1j60sWn9getBj3lan5m6h5WQd8bf64ZBLkhKKzwtVIF743ZA76BT1iR25ZBspEumSEt7MvMBZACrDIVSt7';
const FO_PHONE_NUMBER_ID = '1122823694240869';
const FO_VERIFY_TOKEN = 'field_ops_verify_token_456';

// ── Sheet names ──────────────────────────────────────────────
const SHEET_FAULT   = 'Fault Data';
const SHEET_PD      = 'P-D Data';
const SHEET_FO_SESS = 'FO_Sessions';
const SHEET_FO_LOG  = 'FO_Logs';
const SHEET_FO_HIST = 'FO_History';
const SHEET_REPORT  = 'report';

// ── Admin phones (country code + number, no +) ──────────────
const FO_ADMIN_PHONES = [
  '919999999999',  // replace with real admin number
];

// ── Cluster → WhatsApp Group ID mapping ─────────────────────
const CLUSTER_GROUPS = {
  'VASTRAPUR':   '120363XXXXXXXXX01@g.us',
  'RAILWAYPURA': '120363XXXXXXXXX02@g.us',
  'NARANPURA':   '120363XXXXXXXXX03@g.us',
  'VATVA':       '120363XXXXXXXXX04@g.us',
  'CENTRAL':     '120363XXXXXXXXX05@g.us',
  'SABARMATI':   '120363XXXXXXXXX06@g.us',
  'NAVRANGPURA': '120363XXXXXXXXX07@g.us',
  'GBT':         '120363XXXXXXXXX08@g.us',
  'LC':          '120363XXXXXXXXX09@g.us',
};

// ── Engineer → WhatsApp Group ID mapping ────────────────────
const ENG_GROUPS = {
  'AKSHAY':      '120363XXXXXXXXX10@g.us',
  'BUDDHAJI':    '120363XXXXXXXXX11@g.us',
  'HARISINGH':   '120363XXXXXXXXX12@g.us',
  'ISHWARBHAI':  '120363XXXXXXXXX13@g.us',
  // Add all engineers here...
};

// ── Fault Data column indices (0-based) ─────────────────────
// Headers: Exchange Code | Phone Number | Customer Name |
//   Address | BB User ID | Duration | Repeat Count |
//   Contact Number | ENG | CLUSTER | Status
const F = {
  EXCHANGE: 0,
  PHONE:    1,
  CUST:     2,
  ADDR:     3,
  BBID:     4,
  DUR:      5,
  REPEAT:   6,
  CONTACT:  7,
  ENG:      8,
  CLUSTER:  9,
  STATUS:   10,
};

// ── P-D Data column indices (0-based) ───────────────────────
// Headers: Task Type | Exchange Code | Phone Number |
//   Duration | Customer Name | Address | Contact Number |
//   ENG | CLUSTER | Status
const P = {
  TYPE:     0,
  EXCHANGE: 1,
  PHONE:    2,
  DUR:      3,
  CUST:     4,
  ADDR:     5,
  CONTACT:  6,
  ENG:      7,
  CLUSTER:  8,
  STATUS:   9,
};

const STATUS_PENDING = 'pending';
const STATUS_CLEAR   = 'clear';

// ============================================================
//  WEBHOOK — doGet (single, no duplicate)
// ============================================================
function doGet(e) {
  const p = e.parameter;

  // WhatsApp webhook verification
  if (p['hub.mode'] === 'subscribe' && p['hub.verify_token'] === FO_VERIFY_TOKEN) {
    return ContentService.createTextOutput(p['hub.challenge']);
  }

  const action = p['action'] || '';

  if (action === 'getReportData') return handleGetReportData();
  if (action === 'markClear')     return handleMarkClearAPI(p['type'], p['phone']);
  if (action === 'getTaskRows')   return handleGetTaskRows(p['type'] || 'ALL');
  // FIX v4: New action for dashboard "Dispatch All" button
  if (action === 'dispatchAll')   return handleDispatchAllAPI();

  return ContentService.createTextOutput('Verification failed');
}

// ============================================================
//  DASHBOARD API HANDLERS
// ============================================================

function handleGetReportData() {
  try {
    const ss     = SpreadsheetApp.getActiveSpreadsheet();
    const fSheet = ss.getSheetByName(SHEET_FAULT);
    const pSheet = ss.getSheetByName(SHEET_PD);

    if (!fSheet || !pSheet) {
      return jsonResponse({ error: 'Sheets not found: ' + SHEET_FAULT + ' / ' + SHEET_PD });
    }

    const faultData = fSheet.getDataRange().getValues().slice(1)
      .filter(r => r[F.PHONE] && String(r[F.PHONE]).trim());
    const pdData    = pSheet.getDataRange().getValues().slice(1)
      .filter(r => r[P.PHONE] && String(r[P.PHONE]).trim());

    const engMap = {};

    faultData.forEach(row => {
      const eng     = String(row[F.ENG]     || '').trim().toUpperCase();
      const cluster = String(row[F.CLUSTER] || '').trim().toUpperCase();
      const status  = String(row[F.STATUS]  || '').toLowerCase();
      if (!eng) return;
      const key = cluster + '||' + eng;
      if (!engMap[key]) engMap[key] = { cl: cluster, eng, fa: 0, fc: 0, fp: 0, pa: 0, pc: 0, pp: 0 };
      engMap[key].fa++;
      if (status === STATUS_CLEAR)   engMap[key].fc++;
      if (status === STATUS_PENDING) engMap[key].fp++;
    });

    pdData.forEach(row => {
      const eng     = String(row[P.ENG]     || '').trim().toUpperCase();
      const cluster = String(row[P.CLUSTER] || '').trim().toUpperCase();
      const status  = String(row[P.STATUS]  || '').toLowerCase();
      if (!eng) return;
      const key = cluster + '||' + eng;
      if (!engMap[key]) engMap[key] = { cl: cluster, eng, fa: 0, fc: 0, fp: 0, pa: 0, pc: 0, pp: 0 };
      engMap[key].pa++;
      if (status === STATUS_CLEAR)   engMap[key].pc++;
      if (status === STATUS_PENDING) engMap[key].pp++;
    });

    const CLUSTER_ORDER = ['CENTRAL','GBT','NAVRANGPURA','SABARMATI','VASTRAPUR','VATVA','NARANPURA','LC','RAILWAYPURA'];
    const rows = Object.values(engMap).sort((a, b) => {
      const ci = CLUSTER_ORDER.indexOf(a.cl) - CLUSTER_ORDER.indexOf(b.cl);
      return ci !== 0 ? ci : a.eng.localeCompare(b.eng);
    });

    let lastCluster = '';
    rows.forEach(r => {
      if (r.cl === lastCluster) r.cl = '';
      else lastCluster = r.cl;
    });

    return jsonResponse({
      rows,
      fault_rows:          faultData,
      pd_rows:             pdData,
      ts:                  foNow(),
      total_fault_pending: faultData.filter(r => String(r[F.STATUS]).toLowerCase() === STATUS_PENDING).length,
      total_pd_pending:    pdData.filter(r => String(r[P.STATUS]).toLowerCase() === STATUS_PENDING).length,
    });

  } catch (err) {
    Logger.log('handleGetReportData error: ' + err.message);
    return jsonResponse({ error: err.message });
  }
}

function handleMarkClearAPI(taskType, phone) {
  if (!taskType || !phone) {
    return jsonResponse({ error: 'Missing taskType or phone' });
  }
  try {
    const ss        = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = taskType === 'FAULT' ? SHEET_FAULT : SHEET_PD;
    const sheet     = ss.getSheetByName(sheetName);
    if (!sheet) return jsonResponse({ error: 'Sheet not found: ' + sheetName });

    const data     = sheet.getDataRange().getValues();
    const phoneCol = taskType === 'FAULT' ? F.PHONE  : P.PHONE;
    const statCol  = taskType === 'FAULT' ? F.STATUS : P.STATUS;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][phoneCol]).trim() === String(phone).trim()) {
        sheet.getRange(i + 1, statCol + 1).setValue(STATUS_CLEAR);
        foWriteHistory(
          'dashboard', taskType,
          data[i][taskType === 'FAULT' ? F.CLUSTER : P.CLUSTER],
          1, 'mark_clear_dashboard', foNow()
        );
        return jsonResponse({ success: true, phone, row: i + 1 });
      }
    }
    return jsonResponse({ error: 'Record not found: ' + phone });

  } catch (err) {
    Logger.log('handleMarkClearAPI error: ' + err.message);
    return jsonResponse({ error: err.message });
  }
}

function handleGetTaskRows(type) {
  try {
    let rows = [];
    if (type === 'FAULT' || type === 'ALL') {
      rows = rows.concat(getFaultRows().map(r => ({ _type: 'FAULT', row: r })));
    }
    if (type === 'NTC' || type === 'ALL') {
      rows = rows.concat(getPDRows().filter(r => String(r[P.TYPE]).toUpperCase() === 'NTC').map(r => ({ _type: 'NTC', row: r })));
    }
    if (type === 'DISS' || type === 'ALL') {
      rows = rows.concat(getPDRows().filter(r => String(r[P.TYPE]).toUpperCase() === 'DISS').map(r => ({ _type: 'DISS', row: r })));
    }
    return jsonResponse({ rows, count: rows.length });
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

// FIX v4: New API action — called by dashboard "Dispatch All" button
function handleDispatchAllAPI() {
  try {
    const allPending = getAllPendingByEng();
    const engs = Object.keys(allPending);
    if (engs.length === 0) {
      return jsonResponse({ success: true, sent: 0, message: 'No pending tasks found.' });
    }
    let sentCount = 0, noGroupCount = 0;
    engs.forEach(eng => {
      const tasks = allPending[eng];
      const msg   = buildEngMessage(eng, tasks);
      const total = (tasks.fault?.length || 0) + (tasks.ntc?.length || 0) + (tasks.diss?.length || 0);
      const groupId = findEngGroup(eng);
      if (groupId && groupId.includes('@g.us') && !groupId.includes('XXXXXXXXXX')) {
        foSendText(groupId, msg);
        sentCount++;
        foWriteHistory('dashboard_dispatch', 'ALL_DISPATCH', eng, total, 'eng_dispatch_api', foNow());
      } else {
        noGroupCount++;
      }
      Utilities.sleep(300);
    });
    return jsonResponse({
      success:  true,
      sent:     sentCount,
      skipped:  noGroupCount,
      message:  `Dispatched to ${sentCount} engineer(s). Skipped ${noGroupCount} (no group ID).`
    });
  } catch (err) {
    Logger.log('handleDispatchAllAPI error: ' + err.message);
    return jsonResponse({ error: err.message });
  }
}

// FIX v4: jsonResponse — single definition (removed duplicate from old endpoint file)
function jsonResponse(data) {
  const output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ============================================================
//  doPost — WhatsApp incoming webhook
// ============================================================
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.object !== 'whatsapp_business_account') return ContentService.createTextOutput('OK');

    for (const entry of (body.entry || [])) {
      for (const change of (entry.changes || [])) {
        const value = change.value;
        if (!value.messages) continue;
        for (const msg of value.messages) {
          const phone      = msg.from;
          const senderName = value.contacts?.[0]?.profile?.name || 'User';
          let text = '';
          if (msg.type === 'text') {
            text = msg.text.body;
          } else if (msg.type === 'interactive') {
            const ir = msg.interactive;
            text = ir?.button_reply?.id || ir?.list_reply?.id || ir?.button_reply?.title || ir?.list_reply?.title || '';
          } else {
            foSendText(phone, '⚠️ Please send text only.\n\nType *menu* to see options.');
            continue;
          }
          foWriteLog(phone, senderName, text);
          foHandleMessage(phone, senderName, text.trim());
        }
      }
    }
  } catch (err) {
    Logger.log('doPost error: ' + err.message);
  }
  return ContentService.createTextOutput('OK');
}

// ============================================================
//  MAIN MESSAGE HANDLER
// ============================================================
function foHandleMessage(phone, senderName, messageText) {
  const text    = messageText.toLowerCase().trim();
  const session = foGetSession(phone);

  // Universal cancel
  if (['cancel', 'exit', 'quit', 'stop'].includes(text)) {
    foClearSession(phone);
    foSendText(phone, '❌ *Cancelled.*\n\nType *menu* to start again.');
    return;
  }

  // Main menu
  if (['hi', 'hello', 'menu', 'start', 'home', '0'].includes(text)) {
    foClearSession(phone);
    foSendMainMenu(phone, senderName);
    return;
  }

  // Help
  if (text === 'help') {
    foSendHelp(phone);
    return;
  }

  // ── BUTTON 1: ALL FAULT ──────────────────────────────────
  if (text === '1' || text === 'allfault' || text === 'all fault') {
    foClearSession(phone);
    const stats = buildEngWiseStats();
    foSendText(phone,
      `🔴 *All Pending Tasks — Engineer View*\n━━━━━━━━━━━━━━━━━━━\n\n${stats}\n\n` +
      `Reply *yes* → generate & send all messages to respective groups\n` +
      `Reply *preview* → see messages first before sending\n` +
      `Reply *no* → cancel`
    );
    foSetSession(phone, 'allfault_confirm', {});
    return;
  }

  if (session.step === 'allfault_confirm') {
    if (text === 'yes' || text === 'y') {
      foSetSession(phone, 'idle', {});
      dispatchAllEngMessages(phone, false);
      return;
    }
    if (text === 'preview') {
      foSetSession(phone, 'idle', {});
      dispatchAllEngMessages(phone, true);
      return;
    }
    foClearSession(phone);
    foSendText(phone, '❌ Cancelled.\n\nType *menu* to return.');
    return;
  }

  // ── BUTTON 2: ENG REPORT ────────────────────────────────
  // FIX v4: Changed trigger from "report" to "engreport" / "2"
  // "report" was ambiguous — it collided with user expecting sheet report
  if (text === '2' || text === 'engreport' || text === 'eng report') {
    foClearSession(phone);
    foSendEngReport(phone);
    return;
  }

  // ── BUTTON 3: OVERDUE ───────────────────────────────────
  if (text === '3' || text === 'overdue') {
    foClearSession(phone);
    foSendOverdueReport(phone);
    return;
  }

  // ── OPTION 4–6 ──────────────────────────────────────────
  if (text === '4' || text === 'fault') {
    foClearSession(phone);
    foSetSession(phone, 'fault_filter', {});
    foSendClusterFilter(phone, 'FAULT');
    return;
  }
  if (text === '5' || text === 'ntc') {
    foClearSession(phone);
    foSetSession(phone, 'ntc_filter', {});
    foSendClusterFilter(phone, 'NTC');
    return;
  }
  if (text === '6' || text === 'diss') {
    foClearSession(phone);
    foSetSession(phone, 'diss_filter', {});
    foSendClusterFilter(phone, 'DISS');
    return;
  }

  // ── OPTION 7: MARK CLEAR ────────────────────────────────
  if (text === '7' || text === 'clear' || text === 'mark') {
    foClearSession(phone);
    foSetSession(phone, 'mark_clear_type', {});
    foSendButtons(phone,
      '✅ *Mark as Clear*\n━━━━━━━━━━━━━━━━━━━\n\nSelect task type:',
      [
        { id: 'CLEAR_FAULT', title: '🔴 Fault' },
        { id: 'CLEAR_NTC',   title: '🟡 NTC'   },
        { id: 'CLEAR_DISS',  title: '🔵 DISS'  },
      ]
    );
    return;
  }

  // ── OPTION 8: SUMMARY ───────────────────────────────────
  if (text === '8' || text === 'summary') {
    foSendSummary(phone);
    return;
  }

  // ── OPTION 9: SEARCH ────────────────────────────────────
  if (text === '9' || text === 'search') {
    foClearSession(phone);
    foSetSession(phone, 'search_records', {});
    foSendText(phone,
      '🔍 *Search Records*\n━━━━━━━━━━━━━━━━━━━\n\n' +
      'Search by:\n' +
      '• Phone number (e.g. 079-26860394)\n' +
      '• Customer name\n' +
      '• Engineer name\n' +
      '• Cluster (e.g. VASTRAPUR)\n' +
      '• Status: *pending* or *clear*\n\n' +
      'Type your search term:\n\nType *cancel* to quit.'
    );
    return;
  }

  // ── TODAY ACTIVITY ───────────────────────────────────────
  if (text === 'today') {
    foSendTodayReport(phone);
    return;
  }

  // FIX v4: "reportimage" / "ri" → sends formatted text summary
  // "report" alone no longer collides — reserved for WhatsApp bot context
  if (text === 'reportimage' || text === 'ri') {
    foClearSession(phone);
    foSendReportImageViaURL(phone);
    return;
  }

  // ── Session routing ──────────────────────────────────────
  const step = session.step;

  if (['fault_filter', 'ntc_filter', 'diss_filter'].includes(step)) {
    const taskType = step === 'fault_filter' ? 'FAULT' : step === 'ntc_filter' ? 'NTC' : 'DISS';
    handleClusterFilter(phone, messageText, session, taskType);
    return;
  }
  if (['fault_list', 'ntc_list', 'diss_list'].includes(step)) {
    const taskType = step === 'fault_list' ? 'FAULT' : step === 'ntc_list' ? 'NTC' : 'DISS';
    handleTaskList(phone, messageText, session, taskType);
    return;
  }
  if (step === 'confirm_send')        { handleConfirmSend(phone, text, session); return; }
  if (step === 'mark_clear_type')     { handleMarkClearType(phone, text, session); return; }
  if (step === 'mark_clear_phone')    { handleMarkClearPhone(phone, messageText, session); return; }
  if (step === 'mark_clear_confirm')  { handleMarkClearConfirm(phone, text, session); return; }
  if (step === 'search_records')      { handleSearch(phone, messageText); foClearSession(phone); return; }
  if (step === 'send_eng_report')     { handleSendEngReport(phone, session); return; }

  foSendText(phone, '🤖 Command not recognised.\n\nType *menu* to see all options or *help* for guidance.');
}

// ============================================================
//  BUTTON 1: ALL FAULT — ENGINEER-WISE DISPATCH
// ============================================================
function buildEngWiseStats() {
  const allPending = getAllPendingByEng();
  const engs = Object.keys(allPending);
  if (engs.length === 0) return '🎉 No pending tasks right now!';
  let stats = '';
  engs.forEach(eng => {
    const d = allPending[eng];
    const parts = [];
    if (d.fault.length) parts.push(`🔴 ${d.fault.length} Fault`);
    if (d.ntc.length)   parts.push(`🟡 ${d.ntc.length} NTC`);
    if (d.diss.length)  parts.push(`🔵 ${d.diss.length} DISS`);
    stats += `👷 *${eng}*: ${parts.join(' | ')}\n`;
  });
  return stats;
}

function getAllPendingByEng() {
  const faultRows = getFaultRows().filter(r => String(r[F.STATUS]).toLowerCase() === STATUS_PENDING);
  const pdRows    = getPDRows().filter(r => String(r[P.STATUS]).toLowerCase() === STATUS_PENDING);
  const byEng = {};
  faultRows.forEach(row => {
    const eng = String(row[F.ENG] || '').trim().toUpperCase();
    if (!eng) return;
    if (!byEng[eng]) byEng[eng] = { fault: [], ntc: [], diss: [] };
    byEng[eng].fault.push(row);
  });
  pdRows.forEach(row => {
    const eng  = String(row[P.ENG]  || '').trim().toUpperCase();
    const type = String(row[P.TYPE] || '').trim().toUpperCase();
    if (!eng || !type) return;
    if (!byEng[eng]) byEng[eng] = { fault: [], ntc: [], diss: [] };
    if (type === 'NTC')  byEng[eng].ntc.push(row);
    if (type === 'DISS') byEng[eng].diss.push(row);
  });
  return byEng;
}

function buildEngMessage(engName, tasks) {
  const ts = foNow();
  let msg = `📋 *TASK ASSIGNMENT*\n`;
  msg += `━━━━━━━━━━━━━━━━━━━━━━━━\n`;
  msg += `👷 @${engName}\n`;
  msg += `📅 ${ts}\n`;
  msg += `━━━━━━━━━━━━━━━━━━━━━━━━\n`;

  if (tasks.fault && tasks.fault.length > 0) {
    msg += `\n🔴 *Fault*\n─────────────────────────\n`;
    tasks.fault.forEach(row => {
      const exc  = row[F.EXCHANGE] || '';
      const ph   = row[F.PHONE]    || '';
      const cust = row[F.CUST]     || '';
      const addr = cleanAddress(row[F.ADDR] || '');
      const bbid = row[F.BBID]     || '';
      const dur  = parseFloat(row[F.DUR])    || 0;
      const rpt  = parseFloat(row[F.REPEAT]) || 0;
      const cont = row[F.CONTACT]  || '';
      let line = `${exc} ${ph} ${cust} ${addr} ${bbid} ${cont} ${engName}`;
      if (rpt > 0) line += ` [Repeat:${rpt}]`;
      if (dur > 0) line += ` [${dur}d]`;
      msg += `${line}\n\n`;
    });
  }

  if (tasks.diss && tasks.diss.length > 0) {
    msg += `🔵 *Diss*\n─────────────────────────\n`;
    tasks.diss.forEach(row => {
      const type = row[P.TYPE] || 'DISS';
      const exc  = row[P.EXCHANGE] || '';
      const ph   = row[P.PHONE]    || '';
      const cust = row[P.CUST]     || '';
      const addr = cleanAddress(row[P.ADDR] || '');
      const cont = row[P.CONTACT]  || '';
      const dur  = parseFloat(row[P.DUR])    || 0;
      let line = `${type} ${exc} ${ph} ${cust} ${addr} ${cont} ${engName}`;
      if (dur > 0) line += ` [${dur}d]`;
      msg += `${line}\n\n`;
    });
  }

  if (tasks.ntc && tasks.ntc.length > 0) {
    msg += `🟡 *NTC*\n─────────────────────────\n`;
    tasks.ntc.forEach(row => {
      const type = row[P.TYPE] || 'NTC';
      const exc  = row[P.EXCHANGE] || '';
      const ph   = row[P.PHONE]    || '';
      const dur  = parseFloat(row[P.DUR]) || 0;
      const cust = row[P.CUST]     || '';
      const addr = cleanAddress(row[P.ADDR] || '');
      const cont = row[P.CONTACT]  || '';
      msg += `${type}\t${exc}\t${ph}\t${dur||''}\t${cust}\t${addr}\t${cont}\t${engName}\n\n`;
    });
  }

  const total = (tasks.fault?.length || 0) + (tasks.ntc?.length || 0) + (tasks.diss?.length || 0);
  msg += `━━━━━━━━━━━━━━━━━━━━━━━━\n`;
  msg += `📊 Total Tasks: *${total}*\n`;
  msg += `✔️ _SHIVAM TELECOM Field Bot_`;
  return msg;
}

function cleanAddress(addr) {
  if (!addr) return '';
  return addr.split(',').slice(0, 4).map(s => s.trim()).filter(Boolean).join(', ');
}

function findEngGroup(engName) {
  const upper = String(engName).toUpperCase().trim();
  for (const key of Object.keys(ENG_GROUPS)) {
    if (key.toUpperCase() === upper) return ENG_GROUPS[key];
  }
  const firstEng = upper.split('/')[0].trim();
  for (const key of Object.keys(ENG_GROUPS)) {
    if (key.toUpperCase() === firstEng) return ENG_GROUPS[key];
  }
  return null;
}

function dispatchAllEngMessages(adminPhone, previewOnly) {
  const allPending = getAllPendingByEng();
  const engs = Object.keys(allPending);

  if (engs.length === 0) {
    foSendText(adminPhone, '🎉 *No pending tasks found!*\n\nAll clear.\n\nType *menu* to return.');
    return;
  }

  if (previewOnly) {
    foSendText(adminPhone,
      `👁️ *Preview Mode* — messages will NOT be sent to groups yet.\n` +
      `Generating ${engs.length} engineer message(s)...`
    );
  } else {
    foSendText(adminPhone, `📤 *Dispatching tasks to ${engs.length} engineer(s)...*`);
  }

  let sentCount = 0, noGroupCount = 0, previewSent = 0;

  engs.forEach(eng => {
    const tasks = allPending[eng];
    const msg   = buildEngMessage(eng, tasks);
    const total = (tasks.fault?.length || 0) + (tasks.ntc?.length || 0) + (tasks.diss?.length || 0);

    if (previewOnly) {
      if (previewSent < 5) {
        foSendText(adminPhone, `📋 *PREVIEW for ${eng}:*\n━━━\n${msg}`);
        previewSent++;
      }
    } else {
      const groupId = findEngGroup(eng);
      if (groupId && groupId.includes('@g.us') && !groupId.includes('XXXXXXXXXX')) {
        foSendText(groupId, msg);
        sentCount++;
        foWriteHistory(adminPhone, 'ALL_DISPATCH', eng, total, 'eng_dispatch', foNow());
      } else {
        noGroupCount++;
        foSendText(adminPhone, `⚠️ *No group for ${eng}* — message:\n━━━\n${msg}`);
      }
    }
    Utilities.sleep(300);
  });

  if (previewOnly) {
    foSendText(adminPhone,
      `\n━━━━━━━━━━━━━━━━━━━\n` +
      `Preview shown for ${Math.min(previewSent, engs.length)} of ${engs.length} engineers.\n\n` +
      `Type *yes* to send all now | *no* to cancel`
    );
    foSetSession(adminPhone, 'allfault_confirm', {});
  } else {
    let reply = `✅ *Dispatch Complete!*\n━━━━━━━━━━━━━━━━━━━\n\n`;
    reply += `📤 Sent to groups: *${sentCount}* engineer(s)\n`;
    if (noGroupCount > 0) reply += `⚠️ No group configured: *${noGroupCount}* engineer(s)\n`;
    reply += `\nType *menu* to continue.`;
    foSendText(adminPhone, reply);
  }
}

// ============================================================
//  BUTTON 2: ENGINEER REPORT
// ============================================================
function foSendEngReport(phone) {
  const allPending = getAllPendingByEng();
  const engs = Object.keys(allPending).sort();

  if (engs.length === 0) {
    foSendText(phone, '🎉 *All clear! No pending tasks.*\n\nType *menu* to return.');
    return;
  }

  let msg = `📊 *ENGINEER WORKLOAD REPORT*\n`;
  msg += `━━━━━━━━━━━━━━━━━━━━━━━━\n`;
  msg += `📅 ${foNow()}\n`;
  msg += `━━━━━━━━━━━━━━━━━━━━━━━━\n\n`;

  let totalFault = 0, totalNtc = 0, totalDiss = 0;

  engs.forEach(eng => {
    const d     = allPending[eng];
    const fCnt  = d.fault.length;
    const nCnt  = d.ntc.length;
    const dCnt  = d.diss.length;
    const total = fCnt + nCnt + dCnt;
    totalFault += fCnt; totalNtc += nCnt; totalDiss += dCnt;

    const allDurs = [
      ...d.fault.map(r => parseFloat(r[F.DUR]) || 0),
      ...d.ntc.map(r  => parseFloat(r[P.DUR]) || 0),
      ...d.diss.map(r => parseFloat(r[P.DUR]) || 0),
    ];
    const maxDur = allDurs.length ? Math.max(...allDurs) : 0;
    const overdue = allDurs.filter(x => x > 3).length;
    const urgency = overdue > 0 ? '🚨' : total > 5 ? '⚠️' : '✅';

    msg += `${urgency} *${eng}*\n`;
    if (fCnt) msg += `   🔴 Fault: ${fCnt}\n`;
    if (nCnt) msg += `   🟡 NTC:   ${nCnt}\n`;
    if (dCnt) msg += `   🔵 DISS:  ${dCnt}\n`;
    msg += `   📊 Total: ${total}`;
    if (maxDur > 0)  msg += ` | ⏱ Max: ${maxDur}d`;
    if (overdue > 0) msg += ` | 🚨 Overdue: ${overdue}`;
    msg += `\n───────────────────\n`;
  });

  msg += `\n📈 *TOTALS*\n`;
  msg += `🔴 Fault: *${totalFault}* | 🟡 NTC: *${totalNtc}* | 🔵 DISS: *${totalDiss}*\n`;
  msg += `Grand Total: *${totalFault + totalNtc + totalDiss}* pending tasks\n\n`;
  msg += `_Legend: 🚨 Has overdue (>3d) | ⚠️ High load | ✅ Normal_\n\n`;
  msg += `Type *menu* to return.`;
  foSendText(phone, msg);

  Utilities.sleep(500);
  foSendText(phone,
    `📤 Send this report to cluster group(s)?\n\n` +
    `Reply *sendreport* to send to all cluster groups\n` +
    `Type *menu* to skip`
  );
  foSetSession(phone, 'send_eng_report', { report: msg });
}

function handleSendEngReport(phone, session) {
  const { report } = session.data;
  let sent = 0;
  Object.values(CLUSTER_GROUPS).forEach(gid => {
    if (gid && gid.includes('@g.us') && !gid.includes('XXXXXXXXXX')) {
      foSendText(gid, report);
      sent++;
      Utilities.sleep(200);
    }
  });
  foSendText(phone, `✅ Report sent to *${sent}* cluster group(s)!\n\nType *menu* to continue.`);
  foClearSession(phone);
}

// ============================================================
//  BUTTON 3: OVERDUE REPORT
// ============================================================
function foSendOverdueReport(phone) {
  const threshold = 3;
  const faultRows = getFaultRows().filter(r =>
    String(r[F.STATUS]).toLowerCase() === STATUS_PENDING && parseFloat(r[F.DUR] || 0) > threshold
  );
  const pdRows = getPDRows().filter(r =>
    String(r[P.STATUS]).toLowerCase() === STATUS_PENDING && parseFloat(r[P.DUR] || 0) > threshold
  );
  const total = faultRows.length + pdRows.length;

  if (total === 0) {
    foSendText(phone, `✅ *No overdue tasks!*\n\nAll tasks within ${threshold} days.\n\nType *menu* to return.`);
    return;
  }

  let msg = `🚨 *OVERDUE TASKS (>${threshold} days)*\n`;
  msg += `━━━━━━━━━━━━━━━━━━━━━━━━\n📅 ${foNow()}\nTotal overdue: *${total}*\n━━━━━━━━━━━━━━━━━━━━━━━━\n\n`;

  const byEng = {};
  faultRows.forEach(r => {
    const eng = String(r[F.ENG] || 'Unknown').trim().toUpperCase();
    if (!byEng[eng]) byEng[eng] = [];
    byEng[eng].push({ type: 'FAULT', row: r, dur: parseFloat(r[F.DUR] || 0) });
  });
  pdRows.forEach(r => {
    const eng  = String(r[P.ENG]  || 'Unknown').trim().toUpperCase();
    const type = String(r[P.TYPE] || '').trim().toUpperCase();
    if (!byEng[eng]) byEng[eng] = [];
    byEng[eng].push({ type, row: r, dur: parseFloat(r[P.DUR] || 0) });
  });

  Object.keys(byEng).sort((a,b) => {
    const maxA = Math.max(...byEng[a].map(x => x.dur));
    const maxB = Math.max(...byEng[b].map(x => x.dur));
    return maxB - maxA;
  }).forEach(eng => {
    const items = byEng[eng].sort((a,b) => b.dur - a.dur);
    msg += `👷 *${eng}* (${items.length} overdue)\n`;
    items.slice(0, 5).forEach(({ type, row, dur }) => {
      const typeIcon = type === 'FAULT' ? '🔴' : type === 'NTC' ? '🟡' : '🔵';
      const ph   = type === 'FAULT' ? row[F.PHONE]   : row[P.PHONE];
      const cust = type === 'FAULT' ? row[F.CUST]    : row[P.CUST];
      const clst = type === 'FAULT' ? row[F.CLUSTER] : row[P.CLUSTER];
      msg += `   ${typeIcon} ${ph} | ${cust||'—'} | *${dur}d* | ${clst}\n`;
    });
    if (items.length > 5) msg += `   _...and ${items.length - 5} more_\n`;
    msg += `───────────────────\n`;
  });

  msg += `\nType *menu* to return.`;
  foSendText(phone, msg);
}

// ============================================================
//  CLUSTER FILTER & TASK LIST (options 4/5/6)
// ============================================================
function foSendClusterFilter(phone, taskType) {
  const label = taskType === 'FAULT' ? '🔴 Fault' : taskType === 'NTC' ? '🟡 NTC' : '🔵 DISS';
  foSendButtons(phone,
    `${label}\n━━━━━━━━━━━━━━━━━━━\n\nSelect cluster:`,
    [
      { id: `CL_VASTRAPUR_${taskType}`,   title: '📍 Vastrapur'   },
      { id: `CL_RAILWAYPURA_${taskType}`, title: '📍 Railwaypura' },
      { id: `CL_NARANPURA_${taskType}`,   title: '📍 Naranpura'   },
    ]
  );
  foSendText(phone,
    `Or type cluster name:\n• *VATVA*\n• *CENTRAL*\n• *ALL* (all clusters)\n\nType *cancel* to quit.`
  );
}

function handleClusterFilter(phone, messageText, session, taskType) {
  const step = taskType === 'FAULT' ? 'fault_list' : taskType === 'NTC' ? 'ntc_list' : 'diss_list';
  let cluster = messageText.trim().toUpperCase();
  if (cluster.startsWith('CL_')) {
    cluster = cluster.replace('CL_', '').replace('_' + taskType, '');
  }
  const valid = Object.keys(CLUSTER_GROUPS).concat(['ALL']);
  if (!valid.includes(cluster)) {
    foSendText(phone,
      `❌ Unknown cluster: *${messageText.trim()}*\n\nValid: ${Object.keys(CLUSTER_GROUPS).join(', ')}, ALL\n\nType *cancel* to quit.`
    );
    return;
  }
  foSetSession(phone, step, { cluster, taskType, page: 0 });
  sendTaskList(phone, taskType, cluster, 0);
}

function getFilteredRows(taskType, cluster) {
  let rows = [];
  if (taskType === 'FAULT') {
    rows = getFaultRows().filter(r => String(r[F.STATUS]).toLowerCase() === STATUS_PENDING);
    if (cluster !== 'ALL') rows = rows.filter(r => String(r[F.CLUSTER]).toUpperCase() === cluster);
  } else {
    rows = getPDRows().filter(r =>
      String(r[P.TYPE]).toUpperCase() === taskType &&
      String(r[P.STATUS]).toLowerCase() === STATUS_PENDING
    );
    if (cluster !== 'ALL') rows = rows.filter(r => String(r[P.CLUSTER]).toUpperCase() === cluster);
  }
  return rows;
}

function sendTaskList(phone, taskType, cluster, page) {
  const rows = getFilteredRows(taskType, cluster);
  if (rows.length === 0) {
    foSendText(phone,
      `🎉 *No pending ${taskType} tasks* for ${cluster === 'ALL' ? 'all clusters' : cluster}!\n\nType *menu* to return.`
    );
    foClearSession(phone);
    return;
  }
  const pageSize  = 8;
  const totalPages = Math.ceil(rows.length / pageSize);
  const start     = page * pageSize;
  const slice     = rows.slice(start, start + pageSize);
  const label     = taskType === 'FAULT' ? '🔴 Fault' : taskType === 'NTC' ? '🟡 NTC' : '🔵 DISS';

  let msg = `${label} — ${cluster === 'ALL' ? 'All' : cluster}\n`;
  msg += `━━━━━━━━━━━━━━━━━━━\n*${rows.length}* pending (Page ${page + 1}/${totalPages})\n\n`;

  slice.forEach((row, i) => {
    const num  = start + i + 1;
    const ph   = taskType === 'FAULT' ? row[F.PHONE]   : row[P.PHONE];
    const cust = taskType === 'FAULT' ? row[F.CUST]    : row[P.CUST];
    const eng  = taskType === 'FAULT' ? row[F.ENG]     : row[P.ENG];
    const clst = taskType === 'FAULT' ? row[F.CLUSTER] : row[P.CLUSTER];
    const dur  = taskType === 'FAULT' ? row[F.DUR]     : row[P.DUR];
    msg += `*${num}.* 📞 ${ph}\n`;
    msg += `   👤 ${cust||'—'} | 👷 ${eng||'—'}\n`;
    msg += `   ⏱ ${dur||0}d | 📍 ${clst||'—'}\n───────────────────\n`;
  });

  msg += `\nType *number* to select & send to group`;
  if (page < totalPages - 1) msg += `\nType *next* for more`;
  if (page > 0)               msg += `\nType *prev* for previous`;
  msg += `\nType *all* to send ALL to groups\nType *cancel* to quit.`;
  foSendText(phone, msg);
}

function handleTaskList(phone, messageText, session, taskType) {
  const text    = messageText.toLowerCase().trim();
  const cluster = session.data.cluster;
  const page    = session.data.page || 0;

  if (text === 'next') {
    const np = page + 1;
    foSetSession(phone, session.step, { ...session.data, page: np });
    sendTaskList(phone, taskType, cluster, np);
    return;
  }
  if (text === 'prev' && page > 0) {
    const np = page - 1;
    foSetSession(phone, session.step, { ...session.data, page: np });
    sendTaskList(phone, taskType, cluster, np);
    return;
  }
  if (text === 'all') {
    handleSendAll(phone, taskType, cluster, session);
    return;
  }
  const num = parseInt(messageText.trim());
  if (isNaN(num) || num < 1) {
    foSendText(phone, '❌ Invalid. Type a number, *next*, *prev*, *all*, or *cancel*.');
    return;
  }
  const rows = getFilteredRows(taskType, cluster);
  if (num > rows.length) {
    foSendText(phone, `❌ Max is ${rows.length}. Type *cancel* to quit.`);
    return;
  }
  const row        = rows[num - 1];
  const targetClst = taskType === 'FAULT' ? String(row[F.CLUSTER]) : String(row[P.CLUSTER]);
  const preview    = buildSingleTaskMessage(taskType, row, false);
  foSendText(phone,
    `📋 *Preview:*\n━━━━━━━━━━━━━━━━━━━\n\n${preview}\n━━━━━━━━━━━━━━━━━━━\n📍 Cluster: *${targetClst}*\n\n` +
    `Reply *yes* → send to *${targetClst}* group\nReply *no* → cancel`
  );
  foSetSession(phone, 'confirm_send', { taskType, preview, cluster: targetClst, singleRow: row });
}

function handleSendAll(phone, taskType, cluster, session) {
  const byCluster = {};
  const rows = getFilteredRows(taskType, cluster);
  if (rows.length === 0) {
    foSendText(phone, '🎉 No pending tasks found.\n\nType *menu* to return.');
    foClearSession(phone);
    return;
  }
  rows.forEach(row => {
    const cl = (taskType === 'FAULT' ? String(row[F.CLUSTER]) : String(row[P.CLUSTER])).toUpperCase();
    if (!byCluster[cl]) byCluster[cl] = [];
    byCluster[cl].push(row);
  });
  let summary = `📤 *Send All ${taskType === 'FAULT' ? '🔴 Fault' : taskType === 'NTC' ? '🟡 NTC' : '🔵 DISS'}*\n━━━━━━━━━━━━━━━━━━━\n\n`;
  Object.keys(byCluster).forEach(cl => {
    summary += `📍 *${cl}*: ${byCluster[cl].length} tasks\n`;
  });
  summary += `\nTotal: *${rows.length}* tasks\n\nReply *yes* → send all | *no* → cancel`;
  foSendText(phone, summary);
  foSetSession(phone, 'confirm_send', { taskType, cluster, sendAll: true, byCluster });
}

function handleConfirmSend(phone, text, session) {
  if (text !== 'yes' && text !== 'y') {
    foSendText(phone, '❌ Cancelled.\n\nType *menu* to return.');
    foClearSession(phone);
    return;
  }
  const { taskType, sendAll, byCluster, singleRow, cluster } = session.data;
  if (sendAll && byCluster) {
    let sent = 0, fail = 0;
    const ts = foNow();
    Object.keys(byCluster).forEach(cl => {
      const groupId = CLUSTER_GROUPS[cl];
      if (!groupId || groupId.includes('XXXXXXXXXX')) { fail += byCluster[cl].length; return; }
      const msg = buildClusterBulkMessage(taskType, cl, byCluster[cl], ts);
      foSendText(groupId, msg);
      sent += byCluster[cl].length;
      foWriteHistory(phone, taskType, cl, byCluster[cl].length, 'bulk_cluster', ts);
      Utilities.sleep(300);
    });
    let reply = `✅ *Messages Sent!*\n\n📤 Tasks sent: *${sent}*\n`;
    if (fail > 0) reply += `⚠️ Skipped (no group): *${fail}*\n`;
    reply += `\nType *menu* to continue.`;
    foSendText(phone, reply);
  } else if (singleRow) {
    const clusterKey = String(cluster).toUpperCase();
    const groupId    = CLUSTER_GROUPS[clusterKey];
    if (!groupId || groupId.includes('XXXXXXXXXX')) {
      foSendText(phone, `⚠️ No group configured for *${clusterKey}*.\n\nUpdate CLUSTER_GROUPS.\n\nType *menu* to continue.`);
      foClearSession(phone);
      return;
    }
    const fullMsg = buildSingleTaskMessage(taskType, singleRow, true);
    foSendText(groupId, fullMsg);
    foWriteHistory(phone, taskType, clusterKey, 1, 'single', foNow());
    const ph = taskType === 'FAULT' ? String(singleRow[F.PHONE]) : String(singleRow[P.PHONE]);
    foSendText(phone, `✅ *Task sent to ${clusterKey} group!*\n\n📞 Line: ${ph}\n\nType *menu* to continue.`);
  }
  foClearSession(phone);
}

// ============================================================
//  MESSAGE BUILDERS
// ============================================================
function buildSingleTaskMessage(taskType, row, includeHeader) {
  const ts    = foNow();
  const label = taskType === 'FAULT' ? '🔴 FAULT' : taskType === 'NTC' ? '🟡 NTC' : '🔵 DISS';
  let msg = '';
  if (includeHeader) {
    msg += `${label}\n━━━━━━━━━━━━━━━━━━━━━━━━\n📅 ${ts}\n━━━━━━━━━━━━━━━━━━━━━━━━\n\n`;
  }
  if (taskType === 'FAULT') {
    const eng = String(row[F.ENG] || '—').trim().toUpperCase();
    msg += `👷 *${eng}*\n`;
    msg += `📞 ${row[F.PHONE]   || '—'}\n`;
    msg += `👤 ${row[F.CUST]    || '—'}\n`;
    msg += `🏢 Exchange: ${row[F.EXCHANGE] || '—'}\n`;
    msg += `🏠 ${cleanAddress(row[F.ADDR] || '')}\n`;
    msg += `📱 Contact: ${row[F.CONTACT]  || '—'}\n`;
    msg += `🔢 BB ID: ${row[F.BBID]       || '—'}\n`;
    msg += `⏱ Duration: *${row[F.DUR] || 0} days*\n`;
    if (parseFloat(row[F.REPEAT] || 0) > 0) msg += `🔁 Repeat: ${row[F.REPEAT]} times\n`;
    msg += `📍 Cluster: ${row[F.CLUSTER] || '—'}\n`;
  } else {
    const eng  = String(row[P.ENG]  || '—').trim().toUpperCase();
    const type = String(row[P.TYPE] || taskType);
    msg += `👷 *${eng}*\n`;
    msg += `🏷️ Type: ${type}\n`;
    msg += `🏢 Exchange: ${row[P.EXCHANGE] || '—'}\n`;
    msg += `📞 ${row[P.PHONE]   || '—'}\n`;
    msg += `👤 ${row[P.CUST]    || '—'}\n`;
    msg += `🏠 ${cleanAddress(row[P.ADDR] || '')}\n`;
    msg += `📱 Contact: ${row[P.CONTACT]  || '—'}\n`;
    msg += `⏱ Duration: *${row[P.DUR] || 0} days*\n`;
    msg += `📍 Cluster: ${row[P.CLUSTER] || '—'}\n`;
  }
  if (includeHeader) {
    msg += `\n━━━━━━━━━━━━━━━━━━━━━━━━\n✔️ _SHIVAM TELECOM Field Bot_`;
  }
  return msg;
}

function buildClusterBulkMessage(taskType, cluster, rows, ts) {
  const label = taskType === 'FAULT' ? '🔴 FAULT' : taskType === 'NTC' ? '🟡 NTC' : '🔵 DISS';
  const byEng = {};
  rows.forEach(row => {
    const eng = taskType === 'FAULT' ? String(row[F.ENG] || 'UNKNOWN') : String(row[P.ENG] || 'UNKNOWN');
    const key = eng.trim().toUpperCase();
    if (!byEng[key]) byEng[key] = [];
    byEng[key].push(row);
  });

  let msg = `${label} — *TASK LIST*\n━━━━━━━━━━━━━━━━━━━━━━━━\n📅 ${ts}\n📍 Cluster: *${cluster}*\nTotal: *${rows.length}* tasks\n━━━━━━━━━━━━━━━━━━━━━━━━\n`;
  Object.keys(byEng).sort().forEach(eng => {
    msg += `\n👷 *@${eng}*\n─────────────────────────\n`;
    byEng[eng].forEach(row => {
      if (taskType === 'FAULT') {
        const dur = parseFloat(row[F.DUR]) || 0;
        const rpt = parseFloat(row[F.REPEAT]) || 0;
        let line = `📞 ${row[F.PHONE]} ${row[F.CUST]||''} ${row[F.EXCHANGE]||''} ${cleanAddress(row[F.ADDR]||'')} ${row[F.CONTACT]||''}`;
        if (rpt > 0) line += ` [R:${rpt}]`;
        if (dur > 0) line += ` [${dur}d]`;
        msg += `${line}\n\n`;
      } else {
        msg += `${row[P.TYPE]}\t${row[P.EXCHANGE]||''}\t${row[P.PHONE]}\t${parseFloat(row[P.DUR])||''}\t${row[P.CUST]||''}\t${cleanAddress(row[P.ADDR]||'')}\t${row[P.CONTACT]||''}\n\n`;
      }
    });
  });
  msg += `\n━━━━━━━━━━━━━━━━━━━━━━━━\n✔️ _SHIVAM TELECOM Field Bot_`;
  return msg;
}

// ============================================================
//  MARK AS CLEAR
// ============================================================
function handleMarkClearType(phone, text, session) {
  const t = text.toLowerCase();
  if (t.includes('fault') || t === 'clear_fault') {
    foSetSession(phone, 'mark_clear_phone', { taskType: 'FAULT' });
    foSendText(phone, '🔴 *Mark Fault as Clear*\n\nEnter phone number:\n_e.g. 079-26860394_\n\nType *cancel* to quit.');
  } else if (t.includes('ntc') || t === 'clear_ntc') {
    foSetSession(phone, 'mark_clear_phone', { taskType: 'NTC' });
    foSendText(phone, '🟡 *Mark NTC as Clear*\n\nEnter phone number:\n\nType *cancel* to quit.');
  } else if (t.includes('diss') || t === 'clear_diss') {
    foSetSession(phone, 'mark_clear_phone', { taskType: 'DISS' });
    foSendText(phone, '🔵 *Mark DISS as Clear*\n\nEnter phone number:\n\nType *cancel* to quit.');
  } else {
    foSendText(phone, '❌ Please tap one of the buttons.\nType *cancel* to quit.');
  }
}

function handleMarkClearPhone(phone, messageText, session) {
  const phoneNo  = messageText.trim();
  const taskType = session.data.taskType;
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = taskType === 'FAULT' ? SHEET_FAULT : SHEET_PD;
  const sheet     = ss.getSheetByName(sheetName);

  if (!sheet) {
    foSendText(phone, `❌ Sheet "${sheetName}" not found.\n\nType *menu* to return.`);
    foClearSession(phone);
    return;
  }

  const data  = sheet.getDataRange().getValues();
  let found   = null, rowIndex = -1;

  for (let i = 1; i < data.length; i++) {
    const colIdx = taskType === 'FAULT' ? F.PHONE : P.PHONE;
    if (String(data[i][colIdx]).trim() === phoneNo.trim()) {
      found    = data[i];
      rowIndex = i + 1;
      break;
    }
  }

  if (!found) {
    foSendText(phone,
      `❌ No record for: *${phoneNo}*\n\nCheck format (e.g. *079-26860394*)\n\nType *cancel* to quit.`
    );
    return;
  }

  const custName  = taskType === 'FAULT' ? found[F.CUST]    : found[P.CUST];
  const eng       = taskType === 'FAULT' ? found[F.ENG]     : found[P.ENG];
  const clstr     = taskType === 'FAULT' ? found[F.CLUSTER] : found[P.CLUSTER];
  const curStatus = taskType === 'FAULT' ? found[F.STATUS]  : found[P.STATUS];

  foSendText(phone,
    `📋 *Record Found:*\n━━━━━━━━━━━━━━━━━━━\n\n` +
    `📞 Line: *${phoneNo}*\n` +
    `👤 Customer: ${custName || '—'}\n` +
    `👷 Engineer: ${eng || '—'}\n` +
    `📍 Cluster: ${clstr || '—'}\n` +
    `📊 Status: *${curStatus}*\n\n` +
    `━━━━━━━━━━━━━━━━━━━\n` +
    `Mark as *clear*?\n\nReply *yes* to confirm | *no* to cancel`
  );
  foSetSession(phone, 'mark_clear_confirm', { taskType, phoneNo, rowIndex, custName, eng, cluster: clstr });
}

function handleMarkClearConfirm(phone, text, session) {
  if (text !== 'yes' && text !== 'y') {
    foSendText(phone, '❌ Cancelled.\n\nType *menu* to return.');
    foClearSession(phone);
    return;
  }
  const { taskType, phoneNo, rowIndex, custName, eng, cluster } = session.data;
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(taskType === 'FAULT' ? SHEET_FAULT : SHEET_PD);
    const colIdx = (taskType === 'FAULT' ? F.STATUS : P.STATUS) + 1;
    sheet.getRange(rowIndex, colIdx).setValue(STATUS_CLEAR);
    foSendText(phone,
      `✅ *Marked as Clear!*\n\n` +
      `📞 ${phoneNo}\n👤 ${custName||'—'}\n👷 ${eng||'—'}\n📍 ${cluster||'—'}\n\n` +
      `Type *menu* to continue.`
    );
    foWriteHistory(phone, taskType, cluster, 1, 'mark_clear', foNow());
  } catch (err) {
    Logger.log('Mark clear error: ' + err.message);
    foSendText(phone, '❌ Error updating record.\n\nType *menu* to try again.');
  }
  foClearSession(phone);
}

// ============================================================
//  SEARCH
// ============================================================
function handleSearch(phone, query) {
  const q     = query.toLowerCase().trim();
  const fault = getFaultRows();
  const pd    = getPDRows();
  const results = [];
  fault.filter(r => r.some(c => String(c).toLowerCase().includes(q))).forEach(r => results.push({ type: 'FAULT', row: r }));
  pd.filter(r => r.some(c => String(c).toLowerCase().includes(q))).forEach(r => results.push({ type: String(r[P.TYPE]).toUpperCase(), row: r }));

  if (results.length === 0) {
    foSendText(phone, `🔍 No results for "*${query}*"\n\nType *menu* to return.`);
    return;
  }
  const label = { FAULT: '🔴', NTC: '🟡', DISS: '🔵' };
  let msg = `🔍 *${results.length} result(s):*\n━━━━━━━━━━━━━━━━━━━\n\n`;
  results.slice(0, 8).forEach(({ type, row }) => {
    const icon   = label[type] || '⚪';
    const ph     = type === 'FAULT' ? row[F.PHONE]   : row[P.PHONE];
    const cust   = type === 'FAULT' ? row[F.CUST]    : row[P.CUST];
    const eng    = type === 'FAULT' ? row[F.ENG]     : row[P.ENG];
    const clstr  = type === 'FAULT' ? row[F.CLUSTER] : row[P.CLUSTER];
    const status = type === 'FAULT' ? row[F.STATUS]  : row[P.STATUS];
    const sicon  = String(status).toLowerCase() === STATUS_CLEAR ? '✅' : '⏳';
    msg += `${icon} ${sicon} ${ph} | ${cust||'—'}\n   👷 ${eng||'—'} | 📍 ${clstr||'—'}\n───────────────────\n`;
  });
  if (results.length > 8) msg += `\n_...and ${results.length - 8} more._\n`;
  msg += `\nType *menu* to return.`;
  foSendText(phone, msg);
}

// ============================================================
//  SUMMARY
// ============================================================
function foSendSummary(phone) {
  const fault    = getFaultRows();
  const pd       = getPDRows();
  const ntcRows  = pd.filter(r => String(r[P.TYPE]).toUpperCase() === 'NTC');
  const dissRows = pd.filter(r => String(r[P.TYPE]).toUpperCase() === 'DISS');

  const fPend = fault.filter(r => String(r[F.STATUS]).toLowerCase() === STATUS_PENDING).length;
  const fClr  = fault.filter(r => String(r[F.STATUS]).toLowerCase() === STATUS_CLEAR).length;
  const nPend = ntcRows.filter(r => String(r[P.STATUS]).toLowerCase() === STATUS_PENDING).length;
  const nClr  = ntcRows.filter(r => String(r[P.STATUS]).toLowerCase() === STATUS_CLEAR).length;
  const dPend = dissRows.filter(r => String(r[P.STATUS]).toLowerCase() === STATUS_PENDING).length;
  const dClr  = dissRows.filter(r => String(r[P.STATUS]).toLowerCase() === STATUS_CLEAR).length;

  const byCluster = {};
  fault.filter(r => String(r[F.STATUS]).toLowerCase() === STATUS_PENDING).forEach(r => {
    const cl = String(r[F.CLUSTER]).toUpperCase();
    if (!byCluster[cl]) byCluster[cl] = { f: 0, n: 0, d: 0 };
    byCluster[cl].f++;
  });
  pd.filter(r => String(r[P.STATUS]).toLowerCase() === STATUS_PENDING).forEach(r => {
    const cl   = String(r[P.CLUSTER]).toUpperCase();
    const type = String(r[P.TYPE]).toUpperCase();
    if (!byCluster[cl]) byCluster[cl] = { f: 0, n: 0, d: 0 };
    if (type === 'NTC')  byCluster[cl].n++;
    if (type === 'DISS') byCluster[cl].d++;
  });

  let msg = `📊 *Field Ops Summary*\n━━━━━━━━━━━━━━━━━━━\n\n`;
  msg += `🔴 *Fault*  ⏳ ${fPend} pending | ✅ ${fClr} clear\n`;
  msg += `🟡 *NTC*    ⏳ ${nPend} pending | ✅ ${nClr} clear\n`;
  msg += `🔵 *DISS*   ⏳ ${dPend} pending | ✅ ${dClr} clear\n\n`;
  msg += `━━━━━━━━━━━━━━━━━━━\n📍 *Cluster Breakdown (pending):*\n`;
  Object.keys(byCluster).sort().forEach(cl => {
    const c = byCluster[cl];
    if (c.f + c.n + c.d > 0) {
      const parts = [];
      if (c.f) parts.push(`🔴${c.f}`);
      if (c.n) parts.push(`🟡${c.n}`);
      if (c.d) parts.push(`🔵${c.d}`);
      msg += `\n*${cl}*: ${parts.join(' ')}\n`;
    }
  });
  msg += `\nType *menu* to return.`;
  foSendText(phone, msg);
}

// ============================================================
//  TODAY'S ACTIVITY
// ============================================================
function foSendTodayReport(phone) {
  const today   = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'dd/MM/yyyy');
  const history = getHistoryRows().filter(r => String(r[0]).startsWith(today));
  if (history.length === 0) {
    foSendText(phone, `📭 No activity today (${today}).\n\nType *menu* to return.`);
    return;
  }
  let msg = `📅 *Today's Activity (${today})*\n━━━━━━━━━━━━━━━━━━━\n\n`;
  history.forEach(r => {
    const icon = r[1] === 'FAULT' ? '🔴' : r[1] === 'NTC' ? '🟡' : r[1] === 'DISS' ? '🔵' : '📌';
    msg += `${icon} ${r[1]} | 📍 ${r[2]} | ${r[3]} | ${r[4]} task(s)\n`;
  });
  msg += `\nTotal actions: *${history.length}*\n\nType *menu* to return.`;
  foSendText(phone, msg);
}

// ============================================================
//  REPORT SUMMARY (text-based, triggered by "ri" / "reportimage")
// ============================================================
function foSendReportImageViaURL(phone) {
  try {
    const faultPend  = getFaultRows().filter(r => String(r[F.STATUS]).toLowerCase() === STATUS_PENDING).length;
    const pdPend     = getPDRows().filter(r => String(r[P.STATUS]).toLowerCase() === STATUS_PENDING).length;
    const allPending = getAllPendingByEng();
    const engs       = Object.keys(allPending);

    let msg = `📊 *TASK REPORT*\n━━━━━━━━━━━━━━━━━━━━━━━━\n📅 ${foNow()}\n━━━━━━━━━━━━━━━━━━━━━━━━\n\n`;
    msg += `🔴 Fault Pending: *${faultPend}*\n`;
    msg += `🔵 NTC/DISS Pending: *${pdPend}*\n`;
    msg += `👷 Engineers with tasks: *${engs.length}*\n━━━━━━━━━━━━━━━━━━━━━━━━\n\n`;

    const byCluster = {};
    getFaultRows().filter(r => String(r[F.STATUS]).toLowerCase() === STATUS_PENDING).forEach(r => {
      const cl = String(r[F.CLUSTER]).toUpperCase();
      if (!byCluster[cl]) byCluster[cl] = { f: 0, p: 0 };
      byCluster[cl].f++;
    });
    getPDRows().filter(r => String(r[P.STATUS]).toLowerCase() === STATUS_PENDING).forEach(r => {
      const cl = String(r[P.CLUSTER]).toUpperCase();
      if (!byCluster[cl]) byCluster[cl] = { f: 0, p: 0 };
      byCluster[cl].p++;
    });

    msg += `📍 *Cluster Summary:*\n`;
    Object.keys(byCluster).sort().forEach(cl => {
      const c     = byCluster[cl];
      const parts = [];
      if (c.f) parts.push(`🔴${c.f}`);
      if (c.p) parts.push(`🔵${c.p}`);
      msg += `  *${cl}*: ${parts.join(' ')}\n`;
    });

    msg += `\n━━━━━━━━━━━━━━━━━━━━━━━━\n📋 *Engineer Breakdown:*\n`;
    engs.slice(0, 8).forEach(eng => {
      const d     = allPending[eng];
      const parts = [];
      if (d.fault.length) parts.push(`🔴${d.fault.length}`);
      if (d.ntc.length)   parts.push(`🟡${d.ntc.length}`);
      if (d.diss.length)  parts.push(`🔵${d.diss.length}`);
      msg += `  👷 *${eng}*: ${parts.join(' ')}\n`;
    });
    if (engs.length > 8) msg += `  _...and ${engs.length - 8} more_\n`;
    msg += `\n━━━━━━━━━━━━━━━━━━━━━━━━\n`;
    msg += `💡 Full visual report: Open *Field Ops Dashboard* → Report & Send → One-Click Send\n\n`;
    msg += `✔️ _SHIVAM TELECOM Field Bot_`;

    foSendText(phone, msg);
    foWriteHistory(phone, 'REPORT', 'all', 1, 'report_image_cmd', foNow());
  } catch (err) {
    Logger.log('foSendReportImageViaURL error: ' + err.message);
    foSendText(phone, '❌ Error generating report: ' + err.message + '\n\nType *menu* to return.');
  }
}

// ============================================================
//  MENUS & HELP
// ============================================================
function foSendMainMenu(phone, name) {
  const greeting = (name && name !== 'User') ? `Hello *${name}*! 👋\n\n` : '';
  foSendText(phone,
    `🏢 *SHIVAM TELECOM*\n━━━━━━━━━━━━━━━━━━━\n\n` +
    `${greeting}🔧 *Field Operations Bot v4*\n\n` +
    `━━━ 🚀 *QUICK ACTIONS* ━━━\n\n` +
    `1️⃣  *allfault*    – Send all eng-wise task messages to groups\n` +
    `2️⃣  *engreport*   – Engineer workload report\n` +
    `3️⃣  *overdue*     – Tasks pending >3 days\n\n` +
    `━━━ 🔧 *TASK MANAGEMENT* ━━━\n\n` +
    `4️⃣  *fault*   – Assign individual Fault tasks\n` +
    `5️⃣  *ntc*     – Assign individual NTC tasks\n` +
    `6️⃣  *diss*    – Assign individual DISS tasks\n` +
    `7️⃣  *clear*   – Mark task as clear\n\n` +
    `━━━ 📊 *REPORTS* ━━━\n\n` +
    `8️⃣  *summary*      – Overall status\n` +
    `9️⃣  *search*       – Search any record\n` +
    `📊  *ri*           – Task report summary (text)\n` +
    `    *today*        – Today's activity\n\n` +
    `0️⃣  *help* – Detailed guide\n\n` +
    `_Type *cancel* anytime to stop._`
  );
}

function foSendHelp(phone) {
  foSendText(phone,
    `ℹ️ *Help Guide — Field Operations Bot v4*\n━━━━━━━━━━━━━━━━━━━\n\n` +
    `🚀 *1 / allfault — All Tasks (Recommended)*\n` +
    `Creates one message per engineer with ALL pending Fault + NTC + DISS tasks.\n` +
    `Sends directly to each engineer's WhatsApp group.\n\n` +
    `📊 *2 / engreport — Engineer Report*\n` +
    `Workload per engineer: Fault / NTC / DISS counts, max duration, overdue count.\n\n` +
    `🚨 *3 / overdue — Overdue Tasks*\n` +
    `Lists all tasks pending more than 3 days.\n\n` +
    `🔧 *4–6 / fault, ntc, diss — Individual Assign*\n` +
    `Select cluster → browse tasks → select one or send all to group.\n\n` +
    `✅ *7 / clear — Mark Clear*\n` +
    `Enter phone → confirm → updates status to clear in sheet.\n\n` +
    `📊 *ri / reportimage — Report Summary*\n` +
    `Text-based cluster + engineer summary.\n` +
    `For full image: use Field Ops Dashboard → One-Click Send.\n\n` +
    `━━━━━━━━━━━━━━━━━━━\n` +
    `💡 *Tips:*\n` +
    `• Type *cancel* anytime to stop\n` +
    `• Type *menu* to return to main menu\n` +
    `• Status in sheet must be exactly: *pending* or *clear*\n\n` +
    `Type *menu* to return.`
  );
}

// ============================================================
//  WHATSAPP API
// ============================================================
function foSendText(to, text) {
  const url     = `https://graph.facebook.com/v19.0/${FO_PHONE_NUMBER_ID}/messages`;
  const payload = {
    messaging_product: 'whatsapp',
    to,
    type: 'text',
    text: { body: text, preview_url: false }
  };
  foCallAPI(url, payload);
}

function foSendButtons(to, bodyText, buttons) {
  const url     = `https://graph.facebook.com/v19.0/${FO_PHONE_NUMBER_ID}/messages`;
  const payload = {
    messaging_product: 'whatsapp',
    to,
    type: 'interactive',
    interactive: {
      type: 'button',
      body: { text: bodyText },
      action: {
        buttons: buttons.slice(0, 3).map(b => ({
          type: 'reply',
          reply: { id: b.id, title: b.title.substring(0, 20) }
        }))
      }
    }
  };
  foCallAPI(url, payload);
}

function foCallAPI(url, payload) {
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + FO_WHATSAPP_TOKEN },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };
  try {
    const res  = UrlFetchApp.fetch(url, options);
    const code = res.getResponseCode();
    if (code !== 200) Logger.log('WA Error ' + code + ': ' + res.getContentText());
  } catch (err) {
    Logger.log('WA API Error: ' + err.message);
  }
}

// ============================================================
//  SHEET HELPERS
// ============================================================
function getFaultRows() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_FAULT);
  if (!sheet) return [];
  return sheet.getDataRange().getValues().slice(1).filter(r => r[F.PHONE] && String(r[F.PHONE]).trim());
}

function getPDRows() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PD);
  if (!sheet) return [];
  return sheet.getDataRange().getValues().slice(1).filter(r => r[P.PHONE] && String(r[P.PHONE]).trim());
}

function getHistoryRows() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_FO_HIST);
  if (!sheet) return [];
  return sheet.getDataRange().getValues().slice(1).filter(r => r[0] !== '');
}

// ── Sessions ─────────────────────────────────────────────────
function foGetSessionSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_FO_SESS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_FO_SESS);
    sheet.appendRow(['Phone', 'Step', 'Data', 'Updated']);
    sheet.hideSheet();
  }
  return sheet;
}

function foGetSession(phone) {
  const data = foGetSessionSheet().getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(phone)) {
      try { return { step: data[i][1], data: JSON.parse(data[i][2] || '{}') }; }
      catch (_) { return { step: 'idle', data: {} }; }
    }
  }
  return { step: 'idle', data: {} };
}

function foSetSession(phone, step, data) {
  const sheet  = foGetSessionSheet();
  const values = sheet.getDataRange().getValues();
  const ts     = foNow();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(phone)) {
      sheet.getRange(i + 1, 1, 1, 4).setValues([[phone, step, JSON.stringify(data), ts]]);
      return;
    }
  }
  sheet.appendRow([phone, step, JSON.stringify(data), ts]);
}

function foClearSession(phone) { foSetSession(phone, 'idle', {}); }

// ── Logs ─────────────────────────────────────────────────────
function foGetLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_FO_LOG);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_FO_LOG);
    sheet.appendRow(['Timestamp', 'Phone', 'Name', 'Message']);
    sheet.hideSheet();
  }
  return sheet;
}

function foWriteLog(phone, name, text) {
  try { foGetLogSheet().appendRow([foNow(), phone, name, text]); } catch (_) {}
}

// ── History ──────────────────────────────────────────────────
function foGetHistorySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_FO_HIST);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_FO_HIST);
    sheet.appendRow(['Timestamp', 'Task Type', 'Cluster/Eng', 'Action', 'Count', 'Sent By']);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function foWriteHistory(phone, taskType, target, count, action, ts) {
  try { foGetHistorySheet().appendRow([ts, taskType, target, action, count, phone]); } catch (_) {}
}

function foNow() {
  return Utilities.formatDate(new Date(), 'Asia/Kolkata', 'dd/MM/yyyy HH:mm:ss');
}

// ============================================================
//  SCHEDULED TRIGGERS
// ============================================================
function foMorningSummaryTrigger() {
  FO_ADMIN_PHONES.filter(p => !p.includes('9999999999')).forEach(p => foSendSummary(p));
}

function foCleanOldSessions() {
  const sheet  = foGetSessionSheet();
  const data   = sheet.getDataRange().getValues();
  const cutoff = new Date(Date.now() - 2 * 60 * 60 * 1000);
  for (let i = data.length - 1; i >= 1; i--) {
    try {
      if (new Date(data[i][3]) < cutoff && data[i][1] !== 'idle') {
        sheet.getRange(i + 1, 2, 1, 2).setValues([['idle', '{}']]);
      }
    } catch (_) {}
  }
}

// ============================================================
//  SETUP — run once manually after deploy
// ============================================================
function foSetup() {
  foGetSessionSheet();
  foGetLogSheet();
  foGetHistorySheet();

  Logger.log('✅ Field Ops Bot v4.0 ready!');
  Logger.log('═══════════════════════════════════════════════════');
  Logger.log('v4 CHANGES:');
  Logger.log('  • Duplicate function declarations removed');
  Logger.log('    (doGet, handleGetReportData, etc. were duplicated');
  Logger.log('     between main file and ReportDataEndpoint_v3.gs)');
  Logger.log('  • "report" command clarified:');
  Logger.log('    Use "engreport" (or 2) for engineer workload');
  Logger.log('    Use "ri" or "reportimage" for summary report');
  Logger.log('  • handleDispatchAllAPI added for dashboard button');
  Logger.log('═══════════════════════════════════════════════════');
  Logger.log('COLUMN ORDER CHECK:');
  Logger.log('');
  Logger.log('Fault Data (A→K):');
  Logger.log('  A=Exchange Code, B=Phone, C=Customer Name,');
  Logger.log('  D=Address, E=BB User ID, F=Duration,');
  Logger.log('  G=Repeat Count, H=Contact Number,');
  Logger.log('  I=ENG, J=CLUSTER, K=Status');
  Logger.log('');
  Logger.log('P-D Data (A→J):');
  Logger.log('  A=Task Type, B=Exchange Code, C=Phone,');
  Logger.log('  D=Duration, E=Customer Name, F=Address,');
  Logger.log('  G=Contact Number, H=ENG, I=CLUSTER, J=Status');
  Logger.log('');
  Logger.log('Status values must be: "pending" or "clear"');
  Logger.log('');
  Logger.log('DEPLOY STEPS:');
  Logger.log('1. Use THIS FILE ONLY — do NOT also include');
  Logger.log('   ReportDataEndpoint_v3.gs (it causes duplicate errors)');
  Logger.log('2. Fill ENG_GROUPS with your engineer group IDs');
  Logger.log('3. Fill CLUSTER_GROUPS with cluster group IDs');
  Logger.log('4. Update FO_ADMIN_PHONES');
  Logger.log('5. Deploy → New Deployment → Web App');
  Logger.log('   Execute as: Me | Access: Anyone');
  Logger.log('6. Set webhook in Meta Console');
  Logger.log('   Verify Token = field_ops_verify_token_456');
  Logger.log('7. Optional triggers:');
  Logger.log('   foMorningSummaryTrigger → 9AM daily');
  Logger.log('   foCleanOldSessions → every hour');
  Logger.log('═══════════════════════════════════════════════════');
}