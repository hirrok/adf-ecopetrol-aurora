/**
 * Ecopetrol Aurora — Google Apps Script Backend
 * Version: 2.1 — Production Patch
 * Deploy as: Web App → Execute as: Me → Access: Anyone
 *
 * SHEET TABS (existing tabs preserved, one new tab added):
 *   Fleet | LPG | Contact | Prices | BranchStatus | Changelog  ← NEW
 *
 * AFTER PASTING: Run patchSheets() once to add missing columns + Changelog tab.
 * Then redeploy as a new version (Deploy → Manage deployments → New version).
 */

const SHEET_ID = '1YSl6At1Gbcg64yTtJmL134YkWnNf26rFxm_Q88EX4GM';

// ─── MANAGER AUTH ─────────────────────────────────────────────────────────────
// Manager code lives in Script Properties — never in HTML.
// To set/change: run setManagerCode('yournewcode') from the editor.
function validateManager(code) {
  if (!code) return false;
  const stored = PropertiesService.getScriptProperties().getProperty('MANAGER_CODE');
  if (!stored) return true; // first run: no code set yet, allow through
  return String(code).trim() === String(stored).trim();
}

function setManagerCode(newCode) {
  PropertiesService.getScriptProperties().setProperty('MANAGER_CODE', newCode);
  Logger.log('Manager code updated.');
}

// ─── CHANGELOG ────────────────────────────────────────────────────────────────
function logChange(ss, entry) {
  const sheet = ss.getSheetByName('Changelog');
  if (!sheet) return; // silently skip if tab not created yet
  sheet.appendRow([
    new Date().toISOString(),
    entry.user     || '',
    entry.action   || '',
    entry.entity   || '',
    entry.oldValue || '',
    entry.newValue || '',
    entry.reason   || ''
  ]);
}

// ─── doPost ───────────────────────────────────────────────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const type = data.type;
    const ss   = SpreadsheetApp.openById(SHEET_ID);
    const ts   = new Date().toISOString();

    // ── Public form submissions ──────────────────────────────────────────────
    if (type === 'fleet') {
      const sheet = ss.getSheetByName('Fleet');
      sheet.appendRow([
        ts,
        data.company     || '',
        data.contactName || '',
        data.phone       || '',
        data.email       || '',
        data.fleetSize   || '',
        data.productMix  || '',
        data.branch      || '',
        data.message     || '',
        'New',
        Utilities.getUuid(), // RowId — for status updates
        '',                  // StatusUpdatedBy
        ''                   // StatusUpdatedAt
      ]);
      logChange(ss, {
        user: 'Public Form', action: 'SUBMIT_FLEET',
        entity: data.company || 'Fleet inquiry', newValue: 'New'
      });
    }

    else if (type === 'lpg') {
      const sheet = ss.getSheetByName('LPG');
      sheet.appendRow([
        ts,
        data.name         || '',
        data.phone        || '',
        data.address      || '',
        data.cylinderSize || '',
        data.quantity     || '',
        data.branch       || '',
        data.deliveryDate || '',
        data.customerType || '',
        'New',
        Utilities.getUuid(),
        '',
        ''
      ]);
      logChange(ss, {
        user: 'Public Form', action: 'SUBMIT_LPG',
        entity: data.name || 'LPG request', newValue: 'New'
      });
    }

    else if (type === 'contact') {
      const sheet = ss.getSheetByName('Contact');
      sheet.appendRow([
        ts,
        data.name    || '',
        data.phone   || '',
        data.email   || '',
        data.subject || '',
        data.message || '',
        data.branch  || '',
        'New',
        Utilities.getUuid(),
        '',
        ''
      ]);
      logChange(ss, {
        user: 'Public Form', action: 'SUBMIT_CONTACT',
        entity: data.name || 'Contact message', newValue: 'New'
      });
    }

    // ── Admin: update fuel prices (manager-gated) ────────────────────────────
    else if (type === 'updatePrice') {
      if (!validateManager(data.managerCode)) {
        return jsonResponse({ ok: false, error: 'UNAUTHORIZED' }, 403);
      }

      const sheet    = ss.getSheetByName('Prices');
      const existing = sheet.getDataRange().getValues();
      const latest   = existing.length > 1 ? existing[existing.length - 1] : [];
      const oldStr   = `RON91:${latest[1]||'?'} RON95:${latest[2]||'?'} Diesel:${latest[3]||'?'} Kero:${latest[4]||'?'}`;

      sheet.appendRow([
        ts,
        data.ron91    || latest[1] || '',
        data.ron95    || latest[2] || '',
        data.diesel   || latest[3] || '',
        data.kerosene || latest[4] || '',
        data.updatedBy || 'Manager',
        data.reason    || ''         // Reason column (new)
      ]);

      logChange(ss, {
        user:     data.updatedBy || 'Manager',
        action:   'PRICE_UPDATE',
        entity:   'Prices',
        oldValue: oldStr,
        newValue: `RON91:${data.ron91||''} RON95:${data.ron95||''} Diesel:${data.diesel||''} Kero:${data.kerosene||''}`,
        reason:   data.reason || ''
      });
    }

    // ── Admin: toggle branch status ──────────────────────────────────────────
    else if (type === 'updateBranch') {
      const sheet     = ss.getSheetByName('BranchStatus');
      const dataRange = sheet.getDataRange().getValues();
      let found = false;

      for (let i = 1; i < dataRange.length; i++) {
        if (dataRange[i][0] === data.branchName) {
          if (data.status) sheet.getRange(i + 1, 2).setValue(data.status);
          sheet.getRange(i + 1, 3).setValue(data.note || dataRange[i][2] || '');
          sheet.getRange(i + 1, 4).setValue(ts);
          sheet.getRange(i + 1, 5).setValue(data.updatedBy || 'Employee'); // UpdatedBy column (new)
          found = true;
          break;
        }
      }
      if (!found) {
        sheet.appendRow([data.branchName, data.status || 'Open', data.note || '', ts, data.updatedBy || 'Employee']);
      }

      if (data.status) {
        logChange(ss, {
          user:     data.updatedBy || 'Employee',
          action:   'BRANCH_STATUS',
          entity:   data.branchName,
          newValue: data.status,
          reason:   data.note || data.reason || ''
        });
      }
    }

    // ── Admin: update inquiry status (NEW) ───────────────────────────────────
    else if (type === 'updateStatus') {
      const nameMap   = { fleet: 'Fleet', lpg: 'LPG', contact: 'Contact' };
      const sheetName = nameMap[data.sheet];
      if (!sheetName) return jsonResponse({ ok: false, error: 'Unknown sheet: ' + data.sheet }, 400);

      const sheet   = ss.getSheetByName(sheetName);
      const rows    = sheet.getDataRange().getValues();
      const headers = rows[0].map(h => String(h).trim());

      const statusCol   = headers.indexOf('Status')          + 1;
      const byCol       = headers.indexOf('StatusUpdatedBy') + 1;
      const atCol       = headers.indexOf('StatusUpdatedAt') + 1;
      const rowIdCol    = headers.indexOf('RowId');

      if (statusCol === 0) return jsonResponse({ ok: false, error: 'No Status column' }, 400);

      // Find row: prefer _row (sheet row index), fallback to RowId UUID
      let targetRow = null;
      if (data._row && data._row > 1) {
        targetRow = data._row;
      } else if (data.rowId && rowIdCol >= 0) {
        for (let i = 1; i < rows.length; i++) {
          if (rows[i][rowIdCol] === data.rowId) { targetRow = i + 1; break; }
        }
      }

      if (!targetRow) return jsonResponse({ ok: false, error: 'Row not found' }, 404);

      const oldStatus = sheet.getRange(targetRow, statusCol).getValue();
      sheet.getRange(targetRow, statusCol).setValue(data.status);
      if (byCol > 0) sheet.getRange(targetRow, byCol).setValue(data.updatedBy || 'Employee');
      if (atCol > 0) sheet.getRange(targetRow, atCol).setValue(ts);

      logChange(ss, {
        user:     data.updatedBy || 'Employee',
        action:   'STATUS_UPDATE',
        entity:   sheetName + ' · ' + (data.name || 'row ' + targetRow),
        oldValue: String(oldStatus),
        newValue: data.status,
        reason:   data.reason || ''
      });
    }

    else {
      return jsonResponse({ ok: false, error: 'Unknown type: ' + type }, 400);
    }

    return jsonResponse({ ok: true });

  } catch (err) {
    return jsonResponse({ ok: false, error: err.message }, 500);
  }
}

// ─── doGet ────────────────────────────────────────────────────────────────────
function doGet(e) {
  try {
    const type = e.parameter.type;
    const ss   = SpreadsheetApp.openById(SHEET_ID);

    const sheetMap = {
      fleet:     'Fleet',
      lpg:       'LPG',
      contact:   'Contact',
      prices:    'Prices',
      branches:  'BranchStatus',
      changelog: 'Changelog'    // NEW endpoint
    };

    const sheetName = sheetMap[type];
    if (!sheetName) return jsonResponse({ ok: false, error: 'Unknown type' }, 400);

    const sheet = ss.getSheetByName(sheetName);
    if (!sheet)  return jsonResponse({ ok: true, data: [] });

    const rows    = sheet.getDataRange().getValues();
    const headers = rows[0];
    const records = rows.slice(1)
      .filter(row => row[0]) // skip empty rows
      .map((row, i) => {
        const obj = { _row: i + 2 }; // actual sheet row index
        headers.forEach((h, j) => {
          const val = row[j];
          obj[h] = (val instanceof Date) ? val.toISOString() : val;
        });
        return obj;
      })
      .reverse(); // newest first

    return jsonResponse({ ok: true, data: records });

  } catch (err) {
    return jsonResponse({ ok: false, error: err.message }, 500);
  }
}

// ─── RESPONSE HELPER ─────────────────────────────────────────────────────────
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── PATCH EXISTING SHEETS (run once after paste) ────────────────────────────
// Adds missing columns to existing tabs without destroying data.
// Run: patchSheets() from the editor before redeploying.
function patchSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // 1. Add RowId, StatusUpdatedBy, StatusUpdatedAt to Fleet / LPG / Contact
  const inquirySheets = {
    Fleet:   ['Status','RowId','StatusUpdatedBy','StatusUpdatedAt'],
    LPG:     ['Status','RowId','StatusUpdatedBy','StatusUpdatedAt'],
    Contact: ['Status','RowId','StatusUpdatedBy','StatusUpdatedAt']
  };

  Object.entries(inquirySheets).forEach(([name, required]) => {
    const sheet   = ss.getSheetByName(name);
    if (!sheet) { Logger.log('SKIP: ' + name + ' not found'); return; }
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());

    required.forEach(col => {
      if (!headers.includes(col)) {
        const newCol = sheet.getLastColumn() + 1;
        sheet.getRange(1, newCol).setValue(col).setFontWeight('bold');
        Logger.log('Added column "' + col + '" to ' + name);
      }
    });
  });

  // 2. Add Reason column to Prices if missing
  const prices  = ss.getSheetByName('Prices');
  if (prices) {
    const ph = prices.getRange(1, 1, 1, prices.getLastColumn()).getValues()[0].map(h => String(h).trim());
    if (!ph.includes('Reason')) {
      prices.getRange(1, prices.getLastColumn() + 1).setValue('Reason').setFontWeight('bold');
      Logger.log('Added Reason column to Prices');
    }
  }

  // 3. Add UpdatedBy column to BranchStatus if missing
  const branches = ss.getSheetByName('BranchStatus');
  if (branches) {
    const bh = branches.getRange(1, 1, 1, branches.getLastColumn()).getValues()[0].map(h => String(h).trim());
    if (!bh.includes('UpdatedBy')) {
      branches.getRange(1, branches.getLastColumn() + 1).setValue('UpdatedBy').setFontWeight('bold');
      Logger.log('Added UpdatedBy column to BranchStatus');
    }
  }

  // 4. Create Changelog tab
  let changelog = ss.getSheetByName('Changelog');
  if (!changelog) {
    changelog = ss.insertSheet('Changelog');
    const clHeaders = ['Timestamp','User','Action','Entity','OldValue','NewValue','Reason'];
    changelog.getRange(1, 1, 1, clHeaders.length).setValues([clHeaders])
      .setFontWeight('bold')
      .setBackground('#1a3fa0')
      .setFontColor('#ffffff');
    changelog.setFrozenRows(1);
    Logger.log('Created Changelog tab');
  }

  // 5. Set initial manager code if not set
  const existing = PropertiesService.getScriptProperties().getProperty('MANAGER_CODE');
  if (!existing) {
    PropertiesService.getScriptProperties().setProperty('MANAGER_CODE', 'mgr2026');
    Logger.log('Manager code initialized to: mgr2026 — CHANGE THIS before handing to client.');
  }

  Logger.log('patchSheets() complete.');
}

/**
 * LEGACY: kept for reference. Use patchSheets() on existing installs.
 * RUN ONCE on a fresh sheet: setupSheets() → creates all tabs from scratch.
 */
function setupSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  const tabs = {
    Fleet:        ['Timestamp','Company','ContactName','Phone','Email','FleetSize','ProductMix','Branch','Message','Status','RowId','StatusUpdatedBy','StatusUpdatedAt'],
    LPG:          ['Timestamp','Name','Phone','Address','CylinderSize','Quantity','Branch','DeliveryDate','CustomerType','Status','RowId','StatusUpdatedBy','StatusUpdatedAt'],
    Contact:      ['Timestamp','Name','Phone','Email','Subject','Message','Branch','Status','RowId','StatusUpdatedBy','StatusUpdatedAt'],
    Prices:       ['Timestamp','RON91','RON95','Diesel','Kerosene','UpdatedBy','Reason'],
    BranchStatus: ['BranchName','Status','Note','UpdatedAt','UpdatedBy'],
    Changelog:    ['Timestamp','User','Action','Entity','OldValue','NewValue','Reason']
  };

  Object.entries(tabs).forEach(([name, headers]) => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold')
      .setBackground('#1a3fa0')
      .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  });

  const branches = [
    'Maria Aurora — Flagship','Baler NE','Baler Pingit','Dipaculao South',
    'San Luis','Dinadiawan','Detailen — Maria Aurora','Uno — Maria Aurora'
  ];
  const branchSheet = ss.getSheetByName('BranchStatus');
  branches.forEach(b => branchSheet.appendRow([b, 'Open', '', new Date().toISOString(), 'System']));

  const priceSheet = ss.getSheetByName('Prices');
  priceSheet.appendRow([new Date().toISOString(), '89.50', '94.00', '84.50', '118.50', 'Setup', 'Initial seed']);

  PropertiesService.getScriptProperties().setProperty('MANAGER_CODE', 'mgr2026');
  Logger.log('setupSheets() complete. Change manager code before handoff.');
}
