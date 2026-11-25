/**
 * A'S STORE — Inventory backend (inv-v3)
 * ينحط بنفس مشروعك: ملف جديد (+ Script) اسمه Inventory.
 * ما بتحتاج تمسح إشي من كودك القديم — طلبات الجرد بتروح لهون والباقي لكودك.
 * كل الأسماء هون مبلّشة بـ inv_ / INV_ عشان ما تتعارض مع كودك.
 *
 * الطلبات (كلها GET + JSONP):
 *   ?serial=XXXX                         → تسجيل سكان
 *   ?action=status                       → العدّ الحالي
 *   ?action=preview                      → قائمة الأجهزة غير الممسوحة
 *   ?action=finish&session=Morning|Evening[&reason=...]
 *        → لو في ناقص وما في سبب: يرجّع القائمة (needReason) بدون ما يكتب إشي
 *        → غير هيك: يكتب/يحدّث سطر InventoryChecks (PendingApproval) ويصفّر Scans
 *   ?action=reset                        → تصفير Scans
 */

var INV_CFG = {
  // لو المشروع مفتوح من الشيت (Extensions → Apps Script) خلّيها زي ما هي.
  // لو مشروع مستقل: الصق ID الشيت (من الرابط: /spreadsheets/d/<<ID>>/edit)
  SPREADSHEET_ID: 'PUT_SPREADSHEET_ID_HERE',
  TZ: 'Asia/Amman',
  SH: {
    LIST: 'Missing',          // قائمة الأجهزة الأساسية (المخزون)
    EXCLUDE: 'NotToScan',     // أجهزة داخلية / صيانة — ما بتنحسب
    SCANS: 'Scans',
    CHECKS: 'InventoryChecks',
    UNSCANNED: 'Unscanned',
    RESULT: 'CountResult',
    SUMMARY: 'Summary'
  }
};
var INV_VERSION = 'inv-v4';

/* ======================= Router ======================= */

// ما في داعي تمسح إشي: هاد الملف بياخد طلبات الجرد بس،
// وأي طلب ثاني بيروح لـ doGet القديم تبعك زي ما هو.
// خلّي ملف Inventory آخر ملف بالقائمة (هيك بينضاف تلقائياً).
var INV_PREV_DOGET = (typeof doGet === 'function') ? doGet : null;

function inv_isMine_(p) {
  var a = String(p.action || '').toLowerCase();
  return p.serial != null || ['status', 'preview', 'finish', 'reset'].indexOf(a) >= 0;
}

globalThis.doGet = function (e) {
  var p = (e && e.parameter) || {};
  if (!inv_isMine_(p) && INV_PREV_DOGET) return INV_PREV_DOGET(e);
  var out;
  try {
    out = inv_route_(p);
  } catch (err) {
    out = { cls: 'bad', text: 'خطأ بالسكربت: ' + (err && err.message ? err.message : err) };
  }
  out.version = INV_VERSION;
  return inv_reply_(out, p.callback);
};

function inv_reply_(obj, cb) {
  var json = JSON.stringify(obj);
  if (cb && /^[A-Za-z_$][\w$]{0,63}$/.test(cb)) {
    return ContentService.createTextOutput(cb + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function inv_route_(p) {
  var a = String(p.action || '').toLowerCase();
  if (!a && p.serial != null) return inv_withLock_(function () { return inv_scan_(p.serial); });
  if (a === 'status') return inv_status_();
  if (a === 'preview') return inv_preview_();
  if (a === 'finish') return inv_withLock_(function () { return inv_finish_(p.session, p.reason); });
  if (a === 'reset') return inv_withLock_(inv_reset_);
  return { cls: 'bad', text: 'طلب غير معروف' };
}

function inv_withLock_(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { cls: 'bad', text: 'السيرفر مشغول — جرّب كمان ثانية', busy: true };
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ======================= Helpers ======================= */

function inv_ss_() {
  if (INV_CFG.SPREADSHEET_ID && INV_CFG.SPREADSHEET_ID.indexOf('PUT_') !== 0) {
    return SpreadsheetApp.openById(INV_CFG.SPREADSHEET_ID);
  }
  var a = SpreadsheetApp.getActive();
  if (!a) throw new Error('حط SPREADSHEET_ID بأول السكربت');
  return a;
}
function inv_sh_(ss, name) {
  var s = ss.getSheetByName(name);
  if (!s) throw new Error('الشيت مش موجود: ' + name);
  return s;
}
function inv_norm_(v) { return String(v == null ? '' : v).toUpperCase().replace(/\s+/g, ''); }
function inv_hdr_(row) { return row.map(function (h) { return String(h == null ? '' : h).trim().toLowerCase(); }); }

// يدوّر على العمود: أول إشي تطابق كامل، بعدين "يحتوي"
function inv_col_(H, keys) {
  var i, k;
  for (k = 0; k < keys.length; k++) for (i = 0; i < H.length; i++) if (H[i] === keys[k]) return i;
  for (k = 0; k < keys.length; k++) for (i = 0; i < H.length; i++) if (H[i] && H[i].indexOf(keys[k]) >= 0) return i;
  return -1;
}
function inv_cell_(row, i) { return i >= 0 && row[i] != null ? String(row[i]).trim() : ''; }

function inv_battery_(v) {
  if (v === '' || v == null) return '';
  var n = Number(v);
  if (!isNaN(n)) return (n <= 1 ? Math.round(n * 100) : Math.round(n)) + '%';
  return String(v).trim();
}

function inv_itemLabel_(it) {
  return [it.model, it.color].filter(function (x) { return x; }).join(' · ') || it.serial;
}

/* ======================= Data ======================= */

function inv_loadList_(ss) {
  var data = inv_sh_(ss, INV_CFG.SH.LIST).getDataRange().getValues();
  var H = inv_hdr_(data[0] || []);
  var c = {
    serial: inv_col_(H, ['serial number', 'serial']),
    model: inv_col_(H, ['model']),
    color: inv_col_(H, ['color']),
    storage: inv_col_(H, ['storage']),
    battery: inv_col_(H, ['battery health', 'battery']),
    notes: inv_col_(H, ['notes']),
    extra: inv_col_(H, ['extra note', 'extra'])
  };
  if (c.serial < 0) throw new Error('ما لقيت عمود Serial بشيت ' + INV_CFG.SH.LIST);

  var excluded = inv_loadExcluded_(ss);
  var map = {}, keys = [], skipped = [];
  function cell_has_(row) { return row.some(function (v) { return String(v == null ? '' : v).trim() !== ''; }); }
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var key = inv_norm_(row[c.serial]);
    var rawSerial = String(row[c.serial] == null ? '' : row[c.serial]).trim();
    var why = '';
    if (!key) why = 'بدون سيريال';
    else if (excluded[key]) why = 'NotToScan (داخلي/صيانة)';
    else if (map[key]) why = 'سيريال مكرر';
    if (why) {
      if (key || cell_has_(row)) skipped.push({ serial: rawSerial, model: inv_cell_(row, c.model), why: why });
      continue;
    }
    map[key] = {
      serial: inv_cell_(row, c.serial),
      model: inv_cell_(row, c.model),
      color: inv_cell_(row, c.color),
      storage: inv_cell_(row, c.storage),
      battery: inv_battery_(c.battery >= 0 ? row[c.battery] : ''),
      notes: inv_cell_(row, c.notes),
      extra: inv_cell_(row, c.extra)
    };
    keys.push(key);
  }
  return { map: map, keys: keys, excluded: excluded, skipped: skipped };
}

function inv_loadExcluded_(ss) {
  var out = {};
  var s = ss.getSheetByName(INV_CFG.SH.EXCLUDE);
  if (!s) return out;
  var data = s.getDataRange().getValues();
  var H = inv_hdr_(data[0] || []);
  var cols = [inv_col_(H, ['serial']), inv_col_(H, ['unit id'])];
  for (var r = 1; r < data.length; r++) {
    cols.forEach(function (i) {
      var k = i >= 0 ? inv_norm_(data[r][i]) : '';
      if (k) out[k] = true;
    });
  }
  return out;
}

function inv_loadScans_(ss) {
  var s = inv_sh_(ss, INV_CFG.SH.SCANS);
  var last = s.getLastRow();
  var set = {}, list = [];
  if (last >= 2) {
    var vals = s.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < vals.length; i++) {
      var k = inv_norm_(vals[i][0]);
      if (k && !set[k]) { set[k] = true; list.push(k); }
    }
  }
  return { sheet: s, set: set, list: list };
}

// يطابق الكود الممسوح مع القائمة:
// - الكود كامل
// - كل جزء منه (لو QR فيه رابط أو نص)
// - بدون حرف S بالبداية (باركود علب أبل بيبلّش بـ S)
function inv_match_(raw, L) {
  var full = inv_norm_(raw);
  var tokens = String(raw).toUpperCase().split(/[^A-Z0-9\-]+/).filter(function (t) { return t.length >= 4; });
  var cands = [full].concat(tokens);
  var i, c;
  for (i = 0; i < cands.length; i++) {
    c = cands[i];
    if (L.map[c]) return { type: 'ok', key: c };
    if (c.charAt(0) === 'S' && L.map[c.slice(1)]) return { type: 'ok', key: c.slice(1) };
  }
  for (i = 0; i < cands.length; i++) {
    c = cands[i];
    if (L.excluded[c]) return { type: 'excluded', key: c };
    if (c.charAt(0) === 'S' && L.excluded[c.slice(1)]) return { type: 'excluded', key: c.slice(1) };
  }
  return { type: 'unknown', key: full };
}

function inv_compute_(ss) {
  var L = inv_loadList_(ss);
  var S = inv_loadScans_(ss);
  var items = [], unknown = [], scanned = 0;
  L.keys.forEach(function (k) {
    if (S.set[k]) scanned++;
    else { var it = L.map[k]; it.key = k; items.push(it); }
  });
  S.list.forEach(function (k) { if (!L.map[k] && !L.excluded[k]) unknown.push(k); });
  return { total: L.keys.length, scanned: scanned, missing: items.length, items: items, unknown: unknown, skipped: L.skipped };
}

/* ======================= Actions ======================= */

function inv_scan_(raw) {
  var rawS = String(raw == null ? '' : raw).trim();
  if (!rawS) return { cls: 'bad', text: 'كود فاضي' };
  var ss = inv_ss_();
  var L = inv_loadList_(ss);
  var S = inv_loadScans_(ss);
  var m = inv_match_(rawS, L);
  var scanned = S.list.filter(function (k) { return L.map[k]; }).length;
  var base = { total: L.keys.length, serial: m.key };

  if (m.type === 'excluded') {
    base.cls = 'dup'; base.scanned = scanned;
    base.text = '🔧 ' + m.key + ' — جهاز داخلي/صيانة (مش محسوب)';
    return base;
  }
  if (m.type === 'unknown') {
    if (!S.set[m.key]) S.sheet.appendRow(["'" + rawS.slice(0, 200), new Date()]);
    base.cls = 'bad'; base.unknown = true; base.scanned = scanned;
    base.text = '❓ ' + rawS.slice(0, 60) + ' — مش موجود بقائمة الأجهزة';
    return base;
  }
  var it = L.map[m.key];
  if (S.set[m.key]) {
    base.cls = 'dup'; base.scanned = scanned;
    base.text = '↺ ' + inv_itemLabel_(it) + ' — ممسوح قبل';
    return base;
  }
  S.sheet.appendRow(["'" + m.key, new Date()]);
  base.cls = 'ok'; base.scanned = scanned + 1;
  base.text = '✓ ' + inv_itemLabel_(it) + ' · ' + it.serial;
  return base;
}

function inv_status_() {
  var r = inv_compute_(inv_ss_());
  return { cls: 'ok', text: 'متصل ✓', total: r.total, scanned: r.scanned, missing: r.missing };
}

function inv_preview_() {
  var r = inv_compute_(inv_ss_());
  r.cls = r.missing ? 'dup' : 'ok';
  r.text = r.missing ? ('ناقص ' + r.missing + ' من ' + r.total) : 'كل الأجهزة ممسوحة ✓';
  return r;
}

function inv_finish_(session, reason) {
  session = (String(session) === 'Evening') ? 'Evening' : 'Morning';
  reason = String(reason == null ? '' : reason).trim().slice(0, 500);
  var ss = inv_ss_();
  var r = inv_compute_(ss);

  if (!r.scanned) return { cls: 'bad', text: 'ما في ولا جهاز ممسوح — امسح أولاً' };
  if (r.missing > 0 && !reason) {
    r.cls = 'dup'; r.needReason = true;
    r.text = 'ناقص ' + r.missing + ' — امسحهم أو اكتب السبب';
    return r;
  }

  var now = new Date();
  // اليوم التجاري بينقطع الساعة 2 بالليل (نفس منطق كودك القديم)
  var cut = (typeof DAY_CUTOFF_HOUR !== 'undefined') ? DAY_CUTOFF_HOUR : 2;
  var ymd = Utilities.formatDate(new Date(now.getTime() - cut * 3600 * 1000), INV_CFG.TZ, 'yyyy-MM-dd');
  var auto = r.missing === 0;   // مطابق = موافقة تلقائية
  var checkId = ymd + '_' + session;
  var businessDate = Utilities.parseDate(ymd, ss.getSpreadsheetTimeZone() || INV_CFG.TZ, 'yyyy-MM-dd');

  // 1) الأهم: سطر الاعتماد بالتطبيق
  var up = inv_upsertCheck_(ss, checkId, {
    checkid: checkId,
    businessdate: businessDate,
    session: session,
    total: r.total,
    scanned: r.scanned,
    missing: r.missing,
    status: auto ? 'Approved' : 'PendingApproval',
    reason: reason,
    approvedby: auto ? 'AUTO' : '',
    approvedat: auto ? now : '',
    createdat: now
  });
  checkId = up.id;

  // 2) جداول التقرير — أي خطأ فيها ما بيوقف الجرد
  var warn = [];
  try { inv_writeUnscanned_(ss, r.items, now); } catch (e) { warn.push('Unscanned: ' + e.message); }
  try { inv_writeCountResult_(ss, r); } catch (e) { warn.push('CountResult: ' + e.message); }
  try { inv_writeSummary_(ss, r, now); } catch (e) { warn.push('Summary: ' + e.message); }

  // 3) تصفير السكانات عشان الجلسة الجاية تبلّش نظيفة
  inv_clearBody_(inv_sh_(ss, INV_CFG.SH.SCANS));
  SpreadsheetApp.flush();

  return {
    cls: 'ok', done: true, checkId: checkId, updated: up.updated,
    total: r.total, scanned: r.scanned, missing: r.missing, warnings: warn,
    text: '✓ انبعت الجرد ' + (session === 'Evening' ? 'المسائي' : 'الصباحي') +
      (auto ? ' — مطابق ✓ انعتمد تلقائياً ' : ' للموافقة — ') + r.scanned + '/' + r.total + (r.missing ? (' · ناقص ' + r.missing) : '') +
      (up.updated ? ' (محاولة ' + up.attempt + ')' : '')
  };
}

function inv_reset_() {
  inv_clearBody_(inv_sh_(inv_ss_(), INV_CFG.SH.SCANS));
  SpreadsheetApp.flush();
  return { cls: 'ok', text: '♻️ تم تصفير السكانات — ابدأ من جديد', scanned: 0 };
}

/* ======================= Writers ======================= */

function inv_upsertCheck_(ss, id, rec) {
  var s = inv_sh_(ss, INV_CFG.SH.CHECKS);
  var data = s.getDataRange().getValues();
  var H = inv_hdr_(data[0]);
  var idCol = H.indexOf('checkid');
  if (idCol < 0) throw new Error('عمود CheckID مش موجود بـ ' + INV_CFG.SH.CHECKS);

  // لو نفس الجلسة انبعتت قبل → سطر جديد بـ _2, _3 (ما بنمسح القديم، والمسؤول بيشوف الكل)
  var used = {};
  for (var i = 1; i < data.length; i++) used[String(data[i][idCol]).trim()] = true;
  var finalId = id, n = 1;
  while (used[finalId]) { n++; finalId = id + '_' + n; }
  rec.checkid = finalId;

  var vals = H.map(function (h) { return rec.hasOwnProperty(h) ? rec[h] : ''; });
  s.appendRow(vals);
  return { updated: n > 1, id: finalId, attempt: n };
}

function inv_clearBody_(s) {
  var last = s.getLastRow(), lc = Math.max(1, s.getLastColumn());
  if (last >= 2) s.getRange(2, 1, last - 1, lc).clearContent();
}

function inv_writeRows_(s, objs) {
  inv_clearBody_(s);
  if (!objs.length) return;
  var H = inv_hdr_(s.getRange(1, 1, 1, Math.max(1, s.getLastColumn())).getValues()[0]);
  var rows = objs.map(function (o) { return H.map(function (h) { return o.hasOwnProperty(h) ? o[h] : ''; }); });
  s.getRange(2, 1, rows.length, H.length).setValues(rows);
}

function inv_writeUnscanned_(ss, items, now) {
  var s = ss.getSheetByName(INV_CFG.SH.UNSCANNED); if (!s) return;
  // نحافظ على السبب/الملاحظة/التاريخ اللي انكتبوا قبل لكل جهاز
  var old = {}, data = s.getDataRange().getValues(), H = inv_hdr_(data[0] || []);
  var cS = H.indexOf('serial'), cR = H.indexOf('reason'), cN = H.indexOf('note'), cD = H.indexOf('invdate');
  for (var i = 1; i < data.length; i++) {
    var k = cS >= 0 ? inv_norm_(data[i][cS]) : '';
    if (k) old[k] = { reason: cR >= 0 ? data[i][cR] : '', note: cN >= 0 ? data[i][cN] : '', invdate: cD >= 0 ? data[i][cD] : '' };
  }
  inv_writeRows_(s, items.map(function (it) {
    var o = old[inv_norm_(it.serial)] || {};
    return { serial: it.serial, model: it.model, color: it.color, storage: it.storage,
      battery: it.battery, notes: it.notes, extranote: it.extra,
      reason: o.reason || '', note: o.note || '', invdate: o.invdate || now };
  }));
}

function inv_writeCountResult_(ss, r) {
  var s = ss.getSheetByName(INV_CFG.SH.RESULT); if (!s) return;
  var rows = r.items.map(function (it) { return { group: 'MISSING', code: it.serial, model: it.model, color: it.color }; })
    .concat(r.unknown.map(function (k) { return { group: 'UNKNOWN', code: k }; }));
  inv_writeRows_(s, rows);
}

function inv_writeSummary_(ss, r, now) {
  var s = ss.getSheetByName(INV_CFG.SH.SUMMARY); if (!s) return;
  inv_writeRows_(s, [{ id: '1', total: r.total, scanned: r.scanned, missing: r.missing,
    updatedat: now, invdate: now, approved: false, approvedby: '', approvedat: '' }]);
}

/* شغّلها مرة من المحرر (Run) عشان تعطي الصلاحيات وتتأكد إنه شايف الشيت */
function invTestSetup() {
  Logger.log(JSON.stringify(inv_status_()));
}
