/**
 * マナビス新松戸校 来校管理システム — サーバ側
 *
 * 要件定義書: docs/来校管理システム要件定義.md
 *
 * 構成
 *   doGet()  … 生徒用フォーム（QRから開く） / 受付画面 を配信する
 *   doPost() … 外部から呼ぶ場合のJSON API（GAS配信ページは google.script.run を使う）
 *
 * 初回のみ setup() を実行してシートを作成する。
 */

var TZ = 'Asia/Tokyo';

/* ===================== 設定 ===================== */

var SH = {
  log:  '来校ログ',
  ros:  '生徒名簿',
  res:  '来校予約',
  conf: '設定'
};

var LOG_COLS = ['記録ID', '対象日', '生徒番号', '氏名', '学年', '状況', '来校日時', '目的',
                'ブース', '次回来校予定日', 'ファイル配布日時', '登録経路', '予約時刻', '更新日時'];
var ROS_COLS = ['生徒番号', '氏名', '氏名カナ', '学年', '高校名'];
var RES_COLS = ['対象日', '生徒番号', '氏名', '学年', '予約時刻'];

var DEFAULTS = {
  '遅刻とみなす遅れ(分)': 60,
  '要電話とする遅れ(分)': 60,
  'ブース範囲': '1-2,3-70,71-90,101-120'
};

function prop_(k) { return PropertiesService.getScriptProperties().getProperty(k); }

function ss_() {
  var id = prop_('SHEET_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActive();
}

function sheet_(name, cols) {
  var ss = ss_();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** 初回セットアップ。シートと既定の設定値を作る。 */
function setup() {
  sheet_(SH.log, LOG_COLS);
  sheet_(SH.ros, ROS_COLS);
  sheet_(SH.res, RES_COLS);
  var c = sheet_(SH.conf, ['キー', '値']);
  var have = {};
  c.getDataRange().getValues().slice(1).forEach(function (r) { have[r[0]] = true; });
  Object.keys(DEFAULTS).forEach(function (k) {
    if (!have[k]) c.appendRow([k, DEFAULTS[k]]);
  });
  ss_().setSpreadsheetTimeZone(TZ);
  return 'セットアップ完了';
}

function conf_() {
  var out = {};
  Object.keys(DEFAULTS).forEach(function (k) { out[k] = DEFAULTS[k]; });
  try {
    sheet_(SH.conf, ['キー', '値']).getDataRange().getValues().slice(1).forEach(function (r) {
      if (r[0] !== '' && r[1] !== '') out[r[0]] = r[1];
    });
  } catch (e) { }
  return out;
}

/* ===================== 小道具 ===================== */

function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function nowStr_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'); }

/** セルの値を 'yyyy-MM-dd' に正規化する */
function dstr_(v) {
  if (v === '' || v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  }
  var s = String(v).trim();
  var m = s.match(/(\d{4})\D(\d{1,2})\D(\d{1,2})/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  return s;
}

/** セルの値を 'HH:mm' に正規化する */
function tstr_(v) {
  if (v === '' || v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, TZ, 'HH:mm');
  }
  if (typeof v === 'number') {                 // Excelのシリアル値（1日=1）
    var mins = Math.round((v % 1) * 1440);
    return ('0' + Math.floor(mins / 60)).slice(-2) + ':' + ('0' + (mins % 60)).slice(-2);
  }
  var m = String(v).match(/(\d{1,2}):(\d{2})/);
  return m ? ('0' + m[1]).slice(-2) + ':' + m[2] : '';
}

/**
 * 保存済みの日時から 'HH:mm' を取り出す。
 * new Date() で再パースすると、スクリプトのタイムゾーン設定によって
 * 表示が数時間ずれることがあるため、文字列から直接取り出す。
 * 空欄や不正値では '' を返す（受付画面が落ちないようにする）。
 */
function hhmm_(v) {
  if (v === '' || v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return isNaN(v.getTime()) ? '' : Utilities.formatDate(v, TZ, 'HH:mm');
  }
  var m = String(v).match(/(\d{1,2}):(\d{2})/);
  if (!m) return '';
  if (Number(m[1]) > 23 || Number(m[2]) > 59) return '';
  return ('0' + m[1]).slice(-2) + ':' + m[2];
}

function toMin_(hhmm) {
  var m = String(hhmm || '').match(/(\d{1,2}):(\d{2})/);
  return m ? Number(m[1]) * 60 + Number(m[2]) : null;
}

function nowMin_() {
  var d = new Date();
  return Number(Utilities.formatDate(d, TZ, 'H')) * 60 + Number(Utilities.formatDate(d, TZ, 'm'));
}

/** ブース番号が許可された範囲に入っているか */
function boothOk_(n) {
  var ranges = String(conf_()['ブース範囲']).split(',');
  for (var i = 0; i < ranges.length; i++) {
    var p = ranges[i].split('-');
    if (n >= Number(p[0]) && n <= Number(p[1])) return true;
  }
  return false;
}

function boothKind_(n) {
  if (n >= 1 && n <= 2) return '音読ブース';
  if (n >= 3 && n <= 70) return '受講ブース';
  if (n >= 71 && n <= 90) return '自習ブース';
  if (n >= 101 && n <= 120) return '自習室';
  return '範囲外';
}

function rows_(name, cols) {
  var sh = sheet_(name, cols);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, cols.length).getValues();
  return vals.map(function (r, i) {
    var o = { _row: i + 2 };
    cols.forEach(function (c, j) { o[c] = r[j]; });
    return o;
  }).filter(function (o) { return String(o[cols[0]]).trim() !== ''; });
}

/* ===================== 配信 ===================== */

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.p === 'staff') {
    if (String(p.k || '') !== String(prop_('TOKEN') || '')) {
      return HtmlService.createHtmlOutput('<p style="font:16px -apple-system;padding:24px">合言葉が違います。</p>');
    }
    return page_('staff', '来校管理 受付', {
      token: String(prop_('TOKEN') || ''),
      execUrl: execUrl_()
    });
  }
  return page_('student', '来校登録');
}

/** デプロイ済みウェブアプリのURL（生徒用QRに使う） */
function execUrl_() {
  try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; }
}

function page_(file, title, vars) {
  var t = HtmlService.createTemplateFromFile(file);
  t.pageKey = issueKey_();
  Object.keys(vars || {}).forEach(function (k) { t[k] = vars[k]; });
  return t.evaluate()
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include_(f) { return HtmlService.createHtmlOutputFromFile(f).getContent(); }

/**
 * ページ鍵。生徒番号の総当たり照会を抑えるために、1ページ読み込みあたりの
 * 照会回数に上限を設ける。
 */
function issueKey_() {
  var k = Utilities.getUuid();
  CacheService.getScriptCache().put('pk_' + k, '0', 21600);
  return k;
}

function spendKey_(k) {
  var c = CacheService.getScriptCache();
  var n = Number(c.get('pk_' + k) || -1);
  if (n < 0) throw new Error('ページを開き直してください');
  if (n >= 80) throw new Error('照会が多すぎます。ページを開き直してください');
  c.put('pk_' + k, String(n + 1), 21600);
}

/* ===================== API ===================== */

/** スタッフは合言葉、生徒はページ鍵で通す */
function staff_(req) {
  return !!(req && req.token && String(req.token) === String(prop_('TOKEN') || ''));
}

function guard_(req) {
  if (staff_(req)) return;
  if (req && req.pageKey) { spendKey_(req.pageKey); return; }
  throw new Error('権限がありません');
}

/** 生徒番号から氏名を引く（F-1）。当日すでに登録済みかも返す（F-6） */
function api_lookup(req) {
  guard_(req);
  var no = String(req.no || '').replace(/\D/g, '');
  if (!no) return { ok: false, error: '生徒番号を入力してください' };

  var hit = null;
  rows_(SH.ros, ROS_COLS).forEach(function (r) {
    if (String(r['生徒番号']).replace(/\D/g, '') === no) hit = r;
  });
  if (!hit) return { ok: false, error: 'その番号の生徒が見つかりません。番号を確認してください' };

  var d = today_(), already = null;
  rows_(SH.log, LOG_COLS).forEach(function (r) {
    if (dstr_(r['対象日']) !== d) return;
    if (String(r['生徒番号']).replace(/\D/g, '') !== no) return;
    // 欠席として登録された生徒が実際に来た場合は、登録済みとみなさない
    if (String(r['状況']) === '欠席' || String(r['状況']) === '無断欠席') return;
    already = r;
  });

  var res = null;
  rows_(SH.res, RES_COLS).forEach(function (r) {
    if (dstr_(r['対象日']) === d && String(r['生徒番号']).replace(/\D/g, '') === no) res = r;
  });

  return {
    ok: true,
    no: no,
    name: String(hit['氏名']),
    grade: String(hit['学年']),
    already: !!already,
    alreadyAt: already ? hhmm_(already['来校日時']) : '',
    reservedAt: res ? tstr_(res['予約時刻']) : ''
  };
}

/** 来校登録（F-5）。送信時刻を来校時刻として記録する */
function api_checkin(req) {
  guard_(req);
  var force = staff_(req) && !!req.force;
  var no = String(req.no || '').replace(/\D/g, '');
  var purpose = String(req.purpose || '');
  var booth = Number(req.booth);
  var next = dstr_(req.nextVisit || '');
  var via = (staff_(req) && req.via === '代理') ? '代理' : '本人';

  if (!no) return { ok: false, error: '生徒番号がありません' };
  if (purpose !== '受講' && purpose !== '自習') return { ok: false, error: '受講か自習を選んでください' };
  if (!booth || !boothOk_(booth)) return { ok: false, error: 'ブース番号が範囲外です（1-2 / 3-70 / 71-90 / 101-120）' };
  if (purpose === '自習' && !next) return { ok: false, error: '次回の来校予定日を入力してください' };

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var d = today_();
    var stu = null;
    rows_(SH.ros, ROS_COLS).forEach(function (r) {
      if (String(r['生徒番号']).replace(/\D/g, '') === no) stu = r;
    });
    if (!stu) return { ok: false, error: 'その番号の生徒が見つかりません' };

    // 当日の既存行を探す。欠席の行が残っていれば、二重にせず上書きする。
    var existing = null, absentRow = null;
    rows_(SH.log, LOG_COLS).forEach(function (r) {
      if (dstr_(r['対象日']) !== d) return;
      if (String(r['生徒番号']).replace(/\D/g, '') !== no) return;
      var stt = String(r['状況']);
      if (stt === '欠席' || stt === '無断欠席') absentRow = r;
      else existing = r;
    });
    if (existing && !force) return { ok: false, dup: true, error: '本日はすでに登録されています' };

    // 予約と突合し、正しい来校時刻で遅刻を判定する（S-3 / S-4）
    var resAt = '';
    rows_(SH.res, RES_COLS).forEach(function (r) {
      if (dstr_(r['対象日']) === d && String(r['生徒番号']).replace(/\D/g, '') === no) resAt = tstr_(r['予約時刻']);
    });
    var status = '来校済';
    var limit = Number(conf_()['遅刻とみなす遅れ(分)']);
    if (resAt && nowMin_() > toMin_(resAt) + limit) status = '遅刻';

    var sh = sheet_(SH.log, LOG_COLS);
    var target = existing || absentRow;
    if (target) {
      // 既存行を書き換える（配布済みの記録は引き継ぐ）
      var row = [String(target['記録ID']), d, no, String(stu['氏名']), String(stu['学年']), status,
                 nowStr_(), purpose, booth, next, target['ファイル配布日時'] || '', via, resAt, nowStr_()];
      sh.getRange(target._row, 1, 1, LOG_COLS.length).setValues([row]);
    } else {
      var id = 'V' + Utilities.formatDate(new Date(), TZ, 'yyyyMMddHHmmss') + '-' + no;
      sh.appendRow([id, d, no, String(stu['氏名']), String(stu['学年']), status, nowStr_(),
                    purpose, booth, next, '', via, resAt, nowStr_()]);
    }

    return { ok: true, name: String(stu['氏名']), status: status, booth: booth,
             boothKind: boothKind_(booth), reservedAt: resAt };
  } finally {
    lock.releaseLock();
  }
}

/** 受付画面の一覧（S-1 / S-2 / S-5 / S-7） */
function api_list(req) {
  if (!staff_(req)) return { ok: false, error: '合言葉が違います' };
  var d = dstr_(req.date || '') || today_();
  var c = conf_();

  var logs = rows_(SH.log, LOG_COLS).filter(function (r) { return dstr_(r['対象日']) === d; });
  var visits = logs.filter(function (r) { return String(r['状況']) === '来校済' || String(r['状況']) === '遅刻'; })
    .map(function (r) {
      return {
        id: String(r['記録ID']), no: String(r['生徒番号']), name: String(r['氏名']),
        grade: String(r['学年']), status: String(r['状況']),
        at: hhmm_(r['来校日時']),
        purpose: String(r['目的']), booth: Number(r['ブース']), boothKind: boothKind_(Number(r['ブース'])),
        next: dstr_(r['次回来校予定日']),
        handed: hhmm_(r['ファイル配布日時']),
        via: String(r['登録経路']), reservedAt: tstr_(r['予約時刻'])
      };
    });
  visits.sort(function (a, b) { return b.at.localeCompare(a.at); });

  var absent = logs.filter(function (r) { return String(r['状況']) === '欠席' || String(r['状況']) === '無断欠席'; })
    .map(function (r) {
      return { id: String(r['記録ID']), no: String(r['生徒番号']), name: String(r['氏名']), status: String(r['状況']) };
    });

  // 未来校の予約者と要電話（S-5）
  var done = {};
  logs.forEach(function (r) { done[String(r['生徒番号']).replace(/\D/g, '')] = true; });
  var wait = Number(c['要電話とする遅れ(分)']);
  var pending = [];
  rows_(SH.res, RES_COLS).forEach(function (r) {
    if (dstr_(r['対象日']) !== d) return;
    var no = String(r['生徒番号']).replace(/\D/g, '');
    if (done[no]) return;
    var at = tstr_(r['予約時刻']);
    var late = at && nowMin_() > toMin_(at) + wait;
    pending.push({ no: no, name: String(r['氏名']), grade: String(r['学年']), at: at, alert: !!late });
  });
  pending.sort(function (a, b) { return String(a.at).localeCompare(String(b.at)); });

  return {
    ok: true, date: d, now: Utilities.formatDate(new Date(), TZ, 'HH:mm'),
    visits: visits, absent: absent, pending: pending,
    nextVisits: visits.filter(function (v) { return v.purpose === '自習' && v.next; })
                      .map(function (v) { return { name: v.name, no: v.no, next: v.next }; }),
    counts: {
      visit: visits.length,
      unhanded: visits.filter(function (v) { return !v.handed; }).length,
      alert: pending.filter(function (p) { return p.alert; }).length,
      reserved: pending.length
    },
    rosterCount: rows_(SH.ros, ROS_COLS).length,
    resCount: rows_(SH.res, RES_COLS).filter(function (r) { return dstr_(r['対象日']) === d; }).length
  };
}

/** ファイル配布の記録（S-2） */
function api_handover(req) {
  if (!staff_(req)) return { ok: false, error: '合言葉が違います' };
  var sh = sheet_(SH.log, LOG_COLS);
  var col = LOG_COLS.indexOf('ファイル配布日時') + 1;
  var hit = 0;
  rows_(SH.log, LOG_COLS).forEach(function (r) {
    if (String(r['記録ID']) === String(req.id)) {
      sh.getRange(r._row, col).setValue(req.undo ? '' : nowStr_());
      sh.getRange(r._row, LOG_COLS.indexOf('更新日時') + 1).setValue(nowStr_());
      hit++;
    }
  });
  return hit ? { ok: true } : { ok: false, error: '記録が見つかりません' };
}

/** 欠席・無断欠席の登録（S-6） */
function api_setStatus(req) {
  if (!staff_(req)) return { ok: false, error: '合言葉が違います' };
  var status = String(req.status || '');
  if (['欠席', '無断欠席', '来校済', '遅刻'].indexOf(status) < 0) return { ok: false, error: '状況が不正です' };
  var no = String(req.no || '').replace(/\D/g, '');
  var d = dstr_(req.date || '') || today_();

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = sheet_(SH.log, LOG_COLS);
    var found = null;
    rows_(SH.log, LOG_COLS).forEach(function (r) {
      if (dstr_(r['対象日']) === d && String(r['生徒番号']).replace(/\D/g, '') === no) found = r;
    });
    if (found) {
      sh.getRange(found._row, LOG_COLS.indexOf('状況') + 1).setValue(status);
      if (status === '欠席' || status === '無断欠席') {
        sh.getRange(found._row, LOG_COLS.indexOf('来校日時') + 1).setValue('');
        sh.getRange(found._row, LOG_COLS.indexOf('ブース') + 1).setValue('');
      }
      sh.getRange(found._row, LOG_COLS.indexOf('更新日時') + 1).setValue(nowStr_());
      return { ok: true };
    }
    var stu = null;
    rows_(SH.ros, ROS_COLS).forEach(function (r) {
      if (String(r['生徒番号']).replace(/\D/g, '') === no) stu = r;
    });
    sh.appendRow(['V' + Utilities.formatDate(new Date(), TZ, 'yyyyMMddHHmmss') + '-' + no, d, no,
                  stu ? String(stu['氏名']) : '', stu ? String(stu['学年']) : '', status,
                  '', '', '', '', '', 'スタッフ', '', nowStr_()]);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/** 記録の取り消し */
function api_remove(req) {
  if (!staff_(req)) return { ok: false, error: '合言葉が違います' };
  var sh = sheet_(SH.log, LOG_COLS);
  var target = null;
  rows_(SH.log, LOG_COLS).forEach(function (r) {
    if (String(r['記録ID']) === String(req.id)) target = r;
  });
  if (!target) return { ok: false, error: '記録が見つかりません' };
  sh.deleteRow(target._row);
  return { ok: true };
}

/**
 * 生徒名簿の取り込み（H-1）。
 * 受付画面がExcelを端末内で読み、必要な5項目だけを送ってくる。
 * 保護者氏名・住所・電話番号・生年月日はサーバに届かない（N-3）。
 */
function api_importRoster(req) {
  if (!staff_(req)) return { ok: false, error: '合言葉が違います' };
  var items = req.items || [];
  if (!items.length) return { ok: false, error: '取り込む行がありません' };
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sh = sheet_(SH.ros, ROS_COLS);
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, ROS_COLS.length).clearContent();
    var vals = items.map(function (it) {
      return [String(it.no || '').replace(/\D/g, ''), it.name || '', it.kana || '', it.grade || '', it.school || ''];
    }).filter(function (r) { return r[0]; });
    if (vals.length) sh.getRange(2, 1, vals.length, ROS_COLS.length).setValues(vals);
    return { ok: true, count: vals.length };
  } finally {
    lock.releaseLock();
  }
}

/** 来校予約の取り込み（H-1）。対象日の分を置き換える */
function api_importReservations(req) {
  if (!staff_(req)) return { ok: false, error: '合言葉が違います' };
  var items = req.items || [];
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sh = sheet_(SH.res, RES_COLS);
    var days = {};
    items.forEach(function (it) { if (it.date) days[dstr_(it.date)] = true; });

    // 同じ対象日の既存行を消してから入れる（二重登録を防ぐ）
    var old = rows_(SH.res, RES_COLS);
    for (var i = old.length - 1; i >= 0; i--) {
      if (days[dstr_(old[i]['対象日'])]) sh.deleteRow(old[i]._row);
    }
    var vals = items.map(function (it) {
      return [dstr_(it.date), String(it.no || '').replace(/\D/g, ''), it.name || '', it.grade || '', tstr_(it.at)];
    }).filter(function (r) { return r[0] && r[1]; });
    if (vals.length) sh.getRange(sh.getLastRow() + 1, 1, vals.length, RES_COLS.length).setValues(vals);
    return { ok: true, count: vals.length, days: Object.keys(days) };
  } finally {
    lock.releaseLock();
  }
}

/** 設定の保存 */
function api_saveConf(req) {
  if (!staff_(req)) return { ok: false, error: '合言葉が違います' };
  var sh = sheet_(SH.conf, ['キー', '値']);
  var rowOf = {};
  sh.getDataRange().getValues().slice(1).forEach(function (r, i) { rowOf[r[0]] = i + 2; });
  Object.keys(req.conf || {}).forEach(function (k) {
    if (rowOf[k]) sh.getRange(rowOf[k], 2).setValue(req.conf[k]);
    else sh.appendRow([k, req.conf[k]]);
  });
  return { ok: true };
}

/* ===================== 外部用JSON API ===================== */

function doPost(e) {
  var out;
  try {
    var req = JSON.parse(e.postData.contents);
    var map = {
      lookup: api_lookup, checkin: api_checkin, list: api_list, handover: api_handover,
      setStatus: api_setStatus, remove: api_remove, importRoster: api_importRoster,
      importReservations: api_importReservations, saveConf: api_saveConf
    };
    var fn = map[req.action];
    out = fn ? fn(req) : { ok: false, error: '不明なaction: ' + req.action };
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}
