// ============================================================
// IMPORT TRANSACTION TRACKER
// Paste this entire file at the bottom of your Code.gs
// Then add these cases to your doGet switch:
//
// case 'getImportTransactions':   return wrap(getImportTransactions(data));
// case 'addImportTransaction':    return wrap(addImportTransaction(data));
// case 'updateImportTransaction': return wrap(updateImportTransaction(data));
// case 'getImportDocs':           return wrap(getImportDocs(data));
// case 'updateImportDocs':        return wrap(updateImportDocs(data));
// case 'getImportPayments':       return wrap(getImportPayments(data));
// case 'addImportPayment':        return wrap(addImportPayment(data));
// case 'getImportCosts':          return wrap(getImportCosts(data));
// case 'addImportCost':           return wrap(addImportCost(data));
// case 'deleteImportCost':        return wrap(deleteImportCost(data));
// case 'getImportTimeline':       return wrap(getImportTimeline(data));
// case 'updateImportTimeline':    return wrap(updateImportTimeline(data));
// case 'addImportRemark':         return wrap(addImportRemark(data));
// case 'getImportRemarks':        return wrap(getImportRemarks(data));
// case 'getImportDocReviews':     return wrap(getImportDocReviews(data));
// case 'addImportDocReview':      return wrap(addImportDocReview(data));
//
// Then run setupImportSheets() once from the editor.
// ============================================================

function setupImportSheets() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheets = {
    ImportTransactions: [
      'id','txn_ref','order_ref','buy_direct','supplier','supplier_country',
      'intermediary','intermediary_price','buy_price','buy_currency','buy_conversion_rate','buy_payment_terms',
      'buyer','buyer_country','sell_price','sell_currency','sell_conversion_rate','sell_payment_terms',
      'material','qty','unit','incoterms','vessel','container','bl',
      'port_loading','port_discharge','etd','eta','stage','rag','assigned_to',
      'notes','docs_total','docs_approved','created_by','created_at','updated_at'
    ],
    ImportDocs: [
      'id','txn_id','name','status','reviewer','party',
      'expected_date','received_date','notes','updated_at'
    ],
    ImportPayments: [
      'id','txn_id','leg','type','amount','currency','party',
      'due_date','paid_date','reference','notes','created_by','created_at'
    ],
    ImportCosts: [
      'id','txn_id','type','amount','currency','notes','created_by','created_at'
    ],
    ImportTimeline: [
      'id','txn_id','stage_data','updated_by','updated_at'
    ],
    ImportRemarks: [
      'id','txn_id','remark','added_by','created_at'
    ],
    ImportDocReviews: [
      'id','txn_id','doc_name','action','round','notes','actor','created_at'
    ]
  };
  Object.keys(sheets).forEach(function(name) {
    var headers = sheets[name];
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length)
        .setBackground('#1e3a5f').setFontColor('#ffffff').setFontWeight('bold');
    }
  });
  return { success: true, message: 'Import sheets created' };
}

// Adds missing columns to ImportTransactions without touching existing data.
// Run once from Apps Script editor: addMissingColumns()
function addMissingColumns() {
  var sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('ImportTransactions');
  var existing = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var toAdd = ['buy_conversion_rate', 'sell_conversion_rate'];
  toAdd.forEach(function(col) {
    if (existing.indexOf(col) === -1) {
      var newCol = sh.getLastColumn() + 1;
      sh.getRange(1, newCol).setValue(col)
        .setBackground('#1e3a5f').setFontColor('#ffffff').setFontWeight('bold');
    }
  });
  return { success: true };
}

// Shared helpers

function getImportSheet(name) {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(name);
}

function impHeaders(sheetName) {
  var sh = getImportSheet(sheetName);
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
}

function impRows(sheetName) {
  var sh = getImportSheet(sheetName);
  if (sh.getLastRow() < 2) return [];
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  return sh.getRange(2, 1, sh.getLastRow()-1, sh.getLastColumn()).getValues().map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) { obj[h] = row[i] !== undefined && row[i] !== null ? String(row[i]) : ''; });
    return obj;
  });
}

function impUid() {
  return 'IMP' + Date.now() + Math.random().toString(36).slice(2,6).toUpperCase();
}

function impNow() {
  return new Date().toISOString();
}

function impAppend(sheetName, headers, obj) {
  var sh = getImportSheet(sheetName);
  sh.appendRow(headers.map(function(h) { return obj[h] !== undefined ? obj[h] : ''; }));
}

function impUpdate(sheetName, id, updates) {
  var sh = getImportSheet(sheetName);
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  if (sh.getLastRow() < 2) return false;
  var data = sh.getRange(2, 1, sh.getLastRow()-1, sh.getLastColumn()).getValues();
  var idIdx = headers.indexOf('id');
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][idIdx]) === String(id)) {
      headers.forEach(function(h, j) {
        if (updates[h] !== undefined) sh.getRange(i+2, j+1).setValue(updates[h]);
      });
      return true;
    }
  }
  return false;
}

// Transactions

function getImportTransactions(data) {
  return { transactions: impRows('ImportTransactions') };
}

function addImportTransaction(data) {
  var id = impUid();
  var now = impNow();
  var txn_ref = data.txn_ref || ('IMP-' + new Date().getFullYear() + '-' + String(Math.floor(Math.random()*9000)+1000));
  var headers = impHeaders('ImportTransactions');
  var row = {};
  headers.forEach(function(h) { row[h] = data[h] || ''; });
  row.id = id;
  row.txn_ref = txn_ref;
  row.docs_total = 0;
  row.docs_approved = 0;
  row.created_by = data.by || '';
  row.created_at = now;
  row.updated_at = now;
  impAppend('ImportTransactions', headers, row);
  initImportDocs(id, now);
  initImportTimeline(id, now, data.by || '');
  return { success: true, id: id, txn_ref: txn_ref };
}

function initImportDocs(txn_id, now) {
  var DOC_TYPES = ['Sales Contract','Invoice','OBL','Draft BL','PSIC','Packing List','COO + Form 9','Form 6','Freight Certificate','Insurance','LC'];
  var headers = impHeaders('ImportDocs');
  DOC_TYPES.forEach(function(name) {
    impAppend('ImportDocs', headers, {
      id: impUid(), txn_id: txn_id, name: name, status: 'Pending',
      reviewer: '', party: '', expected_date: '', received_date: '', notes: '', updated_at: now
    });
  });
}

function initImportTimeline(txn_id, now, by) {
  var STAGES = ['Confirmation','Loading','Shipping','Reaches India','Clearance'];
  var timeline = STAGES.map(function(stage) {
    return { stage: stage, expected_date: '', actual_date: '', responsible: '', delay_reason: '', notes: '' };
  });
  var headers = impHeaders('ImportTimeline');
  impAppend('ImportTimeline', headers, {
    id: impUid(), txn_id: txn_id,
    stage_data: JSON.stringify(timeline),
    updated_by: by, updated_at: now
  });
}

function updateImportTransaction(data) {
  var now = impNow();
  var updates = {};
  Object.keys(data).forEach(function(k) { if (k !== 'by') updates[k] = data[k]; });
  updates.updated_at = now;
  var ok = impUpdate('ImportTransactions', data.id, updates);
  return ok ? { success: true } : { error: 'Transaction not found' };
}

// Documents

function getImportDocs(data) {
  return { docs: impRows('ImportDocs').filter(function(r) { return r.txn_id === data.txn_id; }) };
}

function updateImportDocs(data) {
  var now = impNow();
  var docs;
  try { docs = JSON.parse(data.docs); } catch(e) { return { error: 'Invalid docs JSON' }; }
  var sh = getImportSheet('ImportDocs');
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  if (sh.getLastRow() < 2) return { error: 'No doc rows found' };
  var rows = sh.getRange(2, 1, sh.getLastRow()-1, sh.getLastColumn()).getValues();
  var txnIdIdx = headers.indexOf('txn_id');
  var nameIdx  = headers.indexOf('name');
  var updIdx   = headers.indexOf('updated_at');
  var nameToRowIdx = {};
  rows.forEach(function(row, i) {
    if (String(row[txnIdIdx]) === data.txn_id) {
      nameToRowIdx[String(row[nameIdx])] = i;
    }
  });
  docs.forEach(function(d) {
    var ri = nameToRowIdx[d.name];
    if (ri !== undefined) {
      headers.forEach(function(h, j) {
        if (d[h] !== undefined) sh.getRange(ri+2, j+1).setValue(d[h]);
      });
      if (updIdx >= 0) sh.getRange(ri+2, updIdx+1).setValue(now);
    }
  });
  var approved = docs.filter(function(d) { return d.status === 'Approved'; }).length;
  impUpdate('ImportTransactions', data.txn_id, { docs_total: docs.length, docs_approved: approved, updated_at: now });
  return { success: true };
}

// Payments

function getImportPayments(data) {
  return { payments: impRows('ImportPayments').filter(function(r) { return r.txn_id === data.txn_id; }) };
}

function addImportPayment(data) {
  var id = impUid();
  var now = impNow();
  var headers = impHeaders('ImportPayments');
  var row = {};
  headers.forEach(function(h) { row[h] = data[h] || ''; });
  row.id = id;
  row.created_by = data.by || '';
  row.created_at = now;
  impAppend('ImportPayments', headers, row);
  impUpdate('ImportTransactions', data.txn_id, { updated_at: now });
  return { success: true, id: id, created_at: now };
}

// Costs

function getImportCosts(data) {
  return { costs: impRows('ImportCosts').filter(function(r) { return r.txn_id === data.txn_id; }) };
}

function addImportCost(data) {
  var id = impUid();
  var now = impNow();
  var headers = impHeaders('ImportCosts');
  var row = {};
  headers.forEach(function(h) { row[h] = data[h] || ''; });
  row.id = id;
  row.created_by = data.by || '';
  row.created_at = now;
  impAppend('ImportCosts', headers, row);
  impUpdate('ImportTransactions', data.txn_id, { updated_at: now });
  return { success: true, id: id };
}

function deleteImportCost(data) {
  var sh = getImportSheet('ImportCosts');
  if (sh.getLastRow() < 2) return { error: 'Not found' };
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var idIdx = headers.indexOf('id');
  var rows = sh.getRange(2, 1, sh.getLastRow()-1, sh.getLastColumn()).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][idIdx]) === String(data.cost_id)) {
      sh.deleteRow(i+2);
      return { success: true };
    }
  }
  return { error: 'Cost not found' };
}

// Timeline

function getImportTimeline(data) {
  var rows = impRows('ImportTimeline').filter(function(r) { return r.txn_id === data.txn_id; });
  if (!rows.length) return { timeline: [] };
  try {
    return { timeline: JSON.parse(rows[rows.length-1].stage_data) };
  } catch(e) { return { timeline: [] }; }
}

function updateImportTimeline(data) {
  var now = impNow();
  var sh = getImportSheet('ImportTimeline');
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var txnIdx = headers.indexOf('txn_id');
  var sdIdx  = headers.indexOf('stage_data');
  var ubIdx  = headers.indexOf('updated_by');
  var uaIdx  = headers.indexOf('updated_at');
  if (sh.getLastRow() >= 2) {
    var rows = sh.getRange(2, 1, sh.getLastRow()-1, sh.getLastColumn()).getValues();
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][txnIdx]) === data.txn_id) {
        if (sdIdx >= 0) sh.getRange(i+2, sdIdx+1).setValue(data.timeline);
        if (ubIdx >= 0) sh.getRange(i+2, ubIdx+1).setValue(data.by || '');
        if (uaIdx >= 0) sh.getRange(i+2, uaIdx+1).setValue(now);
        impUpdate('ImportTransactions', data.txn_id, { updated_at: now });
        return { success: true };
      }
    }
  }
  impAppend('ImportTimeline', headers, {
    id: impUid(), txn_id: data.txn_id, stage_data: data.timeline,
    updated_by: data.by || '', updated_at: now
  });
  return { success: true };
}

// Remarks

function getImportRemarks(data) {
  var remarks = impRows('ImportRemarks').filter(function(r) { return r.txn_id === data.txn_id; });
  return { remarks: remarks.reverse() };
}

function addImportRemark(data) {
  var id = impUid();
  var now = impNow();
  var headers = impHeaders('ImportRemarks');
  impAppend('ImportRemarks', headers, {
    id: id, txn_id: data.txn_id, remark: data.remark, added_by: data.by || '', created_at: now
  });
  return { success: true, id: id, created_at: now };
}

// Document Review Cycles

function getImportDocReviews(data) {
  return { reviews: impRows('ImportDocReviews').filter(function(r) { return r.txn_id === data.txn_id; }) };
}

function addImportDocReview(data) {
  var id = impUid();
  var now = impNow();
  var headers = impHeaders('ImportDocReviews');
  impAppend('ImportDocReviews', headers, {
    id: id,
    txn_id:   data.txn_id,
    doc_name: data.doc_name,
    action:   data.action,
    round:    data.round || 1,
    notes:    data.notes || '',
    actor:    data.actor || '',
    created_at: now
  });
  impUpdate('ImportTransactions', data.txn_id, { updated_at: now });
  return { success: true, id: id, created_at: now };
}
