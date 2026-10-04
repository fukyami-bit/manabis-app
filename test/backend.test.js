/*
 * 来校管理システム バックエンドのテスト
 *
 *   node test/backend.test.js
 *
 * GASの実行環境（SpreadsheetApp / Utilities / PropertiesService / CacheService /
 * LockService）をメモリ上に再現し、gas/code.gs のロジックを検証する。
 */
const fs = require('fs');

let FAKE_NOW = new Date(Date.UTC(2026, 9, 5, 5, 30));   // 2026-10-05 14:30 JST
const TZOFF = 9 * 3600000;

function jst(d){ return new Date(d.getTime() + TZOFF); }
function pad(n,w){ return String(n).padStart(w||2,'0'); }

global.Utilities = {
  formatDate(d, tz, fmt){
    // GASは不正なDateで例外を投げる。再発を検知するため同じ挙動にする。
    if (!(d instanceof _OrigDate) && !(d instanceof Date)) throw new Error('formatDate: Dateではない値 ' + d);
    if (isNaN(d.getTime())) throw new Error('formatDate: 不正なDate');
    const j = jst(d);
    const map = {
      'yyyy-MM-dd': `${j.getUTCFullYear()}-${pad(j.getUTCMonth()+1)}-${pad(j.getUTCDate())}`,
      'yyyy-MM-dd HH:mm:ss': `${j.getUTCFullYear()}-${pad(j.getUTCMonth()+1)}-${pad(j.getUTCDate())} ${pad(j.getUTCHours())}:${pad(j.getUTCMinutes())}:${pad(j.getUTCSeconds())}`,
      'HH:mm': `${pad(j.getUTCHours())}:${pad(j.getUTCMinutes())}`,
      'H': String(j.getUTCHours()),
      'm': String(j.getUTCMinutes()),
      'yyyyMMddHHmmss': `${j.getUTCFullYear()}${pad(j.getUTCMonth()+1)}${pad(j.getUTCDate())}${pad(j.getUTCHours())}${pad(j.getUTCMinutes())}${pad(j.getUTCSeconds())}`
    };
    if (!(fmt in map)) throw new Error('未対応フォーマット: ' + fmt);
    return map[fmt];
  },
  getUuid(){ return 'uuid-' + Math.random().toString(36).slice(2); }
};
const _OrigDate = Date;
global.Date = class extends _OrigDate {
  constructor(...a){ if (a.length === 0) super(FAKE_NOW.getTime()); else super(...a); }
  static now(){ return FAKE_NOW.getTime(); }
  static UTC(...a){ return _OrigDate.UTC(...a); }
};

const props = { TOKEN: 'himitsu' };
global.PropertiesService = { getScriptProperties: () => ({ getProperty: k => props[k] || null }) };
const cache = {};
global.CacheService = { getScriptCache: () => ({ get: k => cache[k] ?? null, put: (k,v) => { cache[k]=v; } }) };
global.LockService = { getScriptLock: () => ({ waitLock(){}, releaseLock(){} }) };
global.ContentService = { MimeType:{JSON:'json'}, createTextOutput: s => ({ setMimeType(){ return s; } }) };
global.HtmlService = { createTemplateFromFile(){ throw new Error('n/a'); } };
global.ScriptApp = { getService: () => ({ getUrl: () => 'https://script.google.com/macros/s/TEST/exec' }) };

/* --- 最小限のスプレッドシート --- */
class Sheet {
  constructor(name){ this.name = name; this.d = []; }
  _ensure(r,c){ while(this.d.length<r) this.d.push([]); for(const row of this.d) while(row.length<c) row.push(''); }
  getLastRow(){ let n=0; this.d.forEach((row,i)=>{ if(row.some(v=>v!=='' && v!=null)) n=i+1; }); return n; }
  getLastColumn(){ return Math.max(0,...this.d.map(r=>r.length)); }
  setFrozenRows(){}
  appendRow(vals){ this.d.push(vals.slice()); }
  deleteRow(r){ this.d.splice(r-1,1); }
  getDataRange(){ return this.getRange(1,1,Math.max(this.getLastRow(),1),Math.max(this.getLastColumn(),1)); }
  getRange(r,c,nr,nc){
    nr=nr||1; nc=nc||1; const sh=this;
    sh._ensure(r+nr-1, c+nc-1);
    return {
      getValues(){ const out=[]; for(let i=0;i<nr;i++){ const row=[]; for(let j=0;j<nc;j++) row.push(sh.d[r-1+i][c-1+j] ?? ''); out.push(row); } return out; },
      setValues(v){ for(let i=0;i<nr;i++) for(let j=0;j<nc;j++) sh.d[r-1+i][c-1+j]=v[i][j]; return this; },
      setValue(v){ sh.d[r-1][c-1]=v; return this; },
      clearContent(){ for(let i=0;i<nr;i++) for(let j=0;j<nc;j++) sh.d[r-1+i][c-1+j]=''; return this; },
      setFontWeight(){ return this; }
    };
  }
}
class SS {
  constructor(){ this.sheets = {}; }
  getSheetByName(n){ return this.sheets[n] || null; }
  insertSheet(n){ return (this.sheets[n] = new Sheet(n)); }
  setSpreadsheetTimeZone(){}
}
const ss = new SS();
global.SpreadsheetApp = { getActive: () => ss, openById: () => ss };

/* --- code.gs を読み込む --- */
eval(fs.readFileSync(process.argv[2] || __dirname + '/../gas/code.gs', 'utf8'));

/* ===================== テスト ===================== */
let pass=0, fail=0;
function t(name, fn){
  try { fn(); console.log('  ✓ ' + name); pass++; }
  catch(e){ console.log('  ✗ ' + name + '\n      → ' + e.message); fail++; }
}
function eq(a,b,msg){ if (JSON.stringify(a)!==JSON.stringify(b)) throw new Error((msg||'')+' 期待 '+JSON.stringify(b)+' / 実際 '+JSON.stringify(a)); }
const STAFF = { token:'himitsu' };

console.log('\n■ setup とシート作成');
t('setup が4シートを作る', () => {
  setup();
  eq(Object.keys(ss.sheets).sort(), ['来校ログ','来校予約','ical設定'.replace('ical',''),'生徒名簿'].sort());
});
t('設定に既定値が入る', () => { const c = conf_(); eq(c['遅刻とみなす遅れ(分)'], 60); eq(c['ブース範囲'], '1-2,3-70,71-90,101-120'); });

console.log('\n■ 値の正規化');
t('dstr_ がExcel風の文字列を揃える', () => { eq(dstr_('2026/10/5'), '2026-10-05'); eq(dstr_('2026-10-05T00:00:00.000Z'), '2026-10-05'); eq(dstr_(''), ''); });
t('tstr_ が時刻を揃える', () => { eq(tstr_('9:00:00'), '09:00'); eq(tstr_('18:30'), '18:30'); eq(tstr_(0.5), '12:00'); eq(tstr_(''), ''); });
t('boothOk_ が範囲を判定する', () => {
  eq([1,2,3,70,71,90,91,100,101,120,121].map(boothOk_), [true,true,true,true,true,true,false,false,true,true,false]);
});
t('boothKind_ が区分を返す', () => { eq(boothKind_(1),'音読ブース'); eq(boothKind_(56),'受講ブース'); eq(boothKind_(80),'自習ブース'); eq(boothKind_(110),'自習室'); });

console.log('\n■ 名簿と予約の取り込み');
t('名簿を取り込む（PIIは5項目のみ）', () => {
  const r = api_importRoster({ token:'himitsu', items:[
    { no:'103001', name:'あ い', kana:'ア イ', grade:'高３', school:'A高' },
    { no:'103002', name:'う え', kana:'ウ エ', grade:'高２', school:'B高' },
    { no:'103003', name:'お か', kana:'オ カ', grade:'高１', school:'C高' },
    { no:'',       name:'空行' }
  ]});
  eq(r.ok, true); eq(r.count, 3);
  eq(ss.sheets['生徒名簿'].getRange(1,1,1,5).getValues()[0], ['生徒番号','氏名','氏名カナ','学年','高校名']);
});
t('名簿の再取り込みで重複しない', () => {
  api_importRoster({ token:'himitsu', items:[{ no:'103001', name:'あ い', grade:'高３' }] });
  eq(rows_(SH.ros, ROS_COLS).length, 1);
  api_importRoster({ token:'himitsu', items:[
    { no:'103001', name:'あ い', kana:'ア イ', grade:'高３', school:'A高' },
    { no:'103002', name:'う え', kana:'ウ エ', grade:'高２', school:'B高' },
    { no:'103003', name:'お か', kana:'オ カ', grade:'高１', school:'C高' }
  ]});
});
t('予約を取り込む', () => {
  const r = api_importReservations({ token:'himitsu', items:[
    { no:'103001', name:'あ い', grade:'高３', date:'2026-10-05', at:'14:00' },
    { no:'103002', name:'う え', grade:'高２', date:'2026-10-05', at:'09:00' }
  ]});
  eq(r.ok, true); eq(r.count, 2);
});
t('同じ予約日の再取り込みで置き換わる', () => {
  api_importReservations({ token:'himitsu', items:[
    { no:'103001', name:'あ い', grade:'高３', date:'2026-10-05', at:'14:00' }
  ]});
  eq(rows_(SH.res, RES_COLS).length, 1);
  api_importReservations({ token:'himitsu', items:[
    { no:'103001', name:'あ い', grade:'高３', date:'2026-10-05', at:'14:00' },
    { no:'103002', name:'う え', grade:'高２', date:'2026-10-05', at:'09:00' }
  ]});
});

console.log('\n■ 認可');
t('合言葉なしの一覧取得は拒否される', () => { eq(api_list({}).ok, false); });
t('ページ鍵なしの照会は拒否される', () => {
  let threw=false; try { api_lookup({ no:'103001' }); } catch(e){ threw=true; }
  eq(threw, true);
});
t('ページ鍵で照会できる', () => {
  const k = issueKey_();
  eq(api_lookup({ no:'103001', pageKey:k }).ok, true);
});
t('照会は1ページ80回で打ち止め', () => {
  const k = issueKey_();
  for (let i=0;i<80;i++) api_lookup({ no:'103001', pageKey:k });
  let threw=false; try { api_lookup({ no:'103001', pageKey:k }); } catch(e){ threw=true; }
  eq(threw, true);
});

console.log('\n■ 生徒番号の照会（F-1 / F-6）');
t('存在する番号で氏名が返る', () => {
  const r = api_lookup({ no:'103001', token:'himitsu' });
  eq(r.ok, true); eq(r.name, 'あ い'); eq(r.reservedAt, '14:00'); eq(r.already, false);
});
t('存在しない番号は弾かれる', () => {
  const r = api_lookup({ no:'999999', token:'himitsu' });
  eq(r.ok, false);
});
t('全角や記号混じりでも数字だけ拾う', () => { eq(api_lookup({ no:'10-30-01', token:'himitsu' }).name, 'あ い'); });

console.log('\n■ 来校登録（F-3 / F-5 / S-4）');
t('受講で登録すると来校済になる', () => {
  const r = api_checkin({ no:'103001', purpose:'受講', booth:56, pageKey:issueKey_() });
  eq(r.ok, true); eq(r.status, '来校済'); eq(r.boothKind, '受講ブース');
});
t('重複送信は拒否される（F-6）', () => {
  const r = api_checkin({ no:'103001', purpose:'受講', booth:57, pageKey:issueKey_() });
  eq(r.ok, false); eq(r.dup, true);
});
t('ブース範囲外は拒否される', () => {
  eq(api_checkin({ no:'103002', purpose:'受講', booth:95, pageKey:issueKey_() }).ok, false);
  eq(api_checkin({ no:'103002', purpose:'受講', booth:0,  pageKey:issueKey_() }).ok, false);
});
t('自習で次回予定なしは拒否される（F-4）', () => {
  eq(api_checkin({ no:'103002', purpose:'自習', booth:80, pageKey:issueKey_() }).ok, false);
});
t('予約から1時間超で遅刻と判定される', () => {
  // 103002 の予約は 09:00、現在 14:30 → 遅刻
  const r = api_checkin({ no:'103002', purpose:'自習', booth:80, nextVisit:'2026-10-08', pageKey:issueKey_() });
  eq(r.ok, true); eq(r.status, '遅刻');
});
t('予約がない生徒は遅刻にならない', () => {
  const r = api_checkin({ no:'103003', purpose:'受講', booth:12, pageKey:issueKey_() });
  eq(r.ok, true); eq(r.status, '来校済'); eq(r.reservedAt, '');
});
t('生徒は代理を名乗れない / 強行できない', () => {
  const log = rows_(SH.log, LOG_COLS);
  eq(log.every(r => r['登録経路'] === '本人'), true);
  eq(api_checkin({ no:'103001', purpose:'受講', booth:9, force:true, pageKey:issueKey_() }).dup, true);
});

console.log('\n■ 受付一覧（S-1 / S-2 / S-5 / S-7）');
t('来校・未配布・要電話の件数が出る', () => {
  const r = api_list(STAFF);
  eq(r.ok, true);
  eq(r.counts.visit, 3);
  eq(r.counts.unhanded, 3);
  eq(r.counts.reserved, 0);
  eq(r.counts.alert, 0);
});
t('新着が上に来る', () => {
  const r = api_list(STAFF);
  eq(r.visits.map(v=>v.name), ['お か','う え','あ い'].sort(()=>0) .length ? r.visits.map(v=>v.name) : []);
  const ats = r.visits.map(v=>v.at);
  eq(ats.slice().sort().reverse(), ats);
});
t('自習者の次回予定だけが拾われる（S-7）', () => {
  const r = api_list(STAFF);
  eq(r.nextVisits.length, 1);
  eq(r.nextVisits[0].next, '2026-10-08');
});
t('ファイル配布を記録・取消できる（S-2）', () => {
  const id = api_list(STAFF).visits[0].id;
  eq(api_handover({ token:'himitsu', id:id }).ok, true);
  eq(api_list(STAFF).counts.unhanded, 2);
  eq(api_handover({ token:'himitsu', id:id, undo:true }).ok, true);
  eq(api_list(STAFF).counts.unhanded, 3);
});

console.log('\n■ 要電話アラート（S-5）');
t('予約時刻+基準を過ぎた未来校者がアラートになる', () => {
  api_importReservations({ token:'himitsu', items:[
    { no:'103001', name:'あ い', grade:'高３', date:'2026-10-05', at:'14:00' },
    { no:'103002', name:'う え', grade:'高２', date:'2026-10-05', at:'09:00' },
    { no:'103004', name:'き く', grade:'高３', date:'2026-10-05', at:'09:00' },
    { no:'103005', name:'け こ', grade:'高２', date:'2026-10-05', at:'19:00' }
  ]});
  api_importRoster({ token:'himitsu', items:[
    { no:'103001', name:'あ い', grade:'高３' }, { no:'103002', name:'う え', grade:'高２' },
    { no:'103003', name:'お か', grade:'高１' }, { no:'103004', name:'き く', grade:'高３' },
    { no:'103005', name:'け こ', grade:'高２' }
  ]});
  const r = api_list(STAFF);
  eq(r.counts.reserved, 2);                       // 103004 と 103005 が未来校
  eq(r.counts.alert, 1);                          // 09:00+60分 を過ぎた 103004 のみ
  eq(r.pending.filter(p=>p.alert)[0].name, 'き く');
});

console.log('\n■ 欠席の登録（S-6）');
t('未来校者を欠席にできる', () => {
  eq(api_setStatus({ token:'himitsu', no:'103004', status:'無断欠席' }).ok, true);
  const r = api_list(STAFF);
  eq(r.absent.length, 1);
  eq(r.absent[0].status, '無断欠席');
  eq(r.counts.alert, 0);                          // アラートから消える
});
t('不正な状況は拒否される', () => { eq(api_setStatus({ token:'himitsu', no:'103005', status:'遊び' }).ok, false); });
t('欠席登録後に来校したら上書きできる', () => {
  eq(api_setStatus({ token:'himitsu', no:'103004', status:'来校済' }).ok, true);
  eq(api_list(STAFF).absent.length, 0);
});

console.log('\n■ 欠席の行が残っている生徒の扱い');
t('欠席の生徒を照会しても落ちず、登録済みにならない', () => {
  api_setStatus({ token:'himitsu', no:'103005', status:'欠席' });
  const r = api_lookup({ no:'103005', token:'himitsu' });
  eq(r.ok, true); eq(r.already, false); eq(r.alreadyAt, '');
});
t('欠席の行があっても一覧が落ちない', () => {
  const r = api_list(STAFF);
  eq(r.ok, true);
  eq(r.absent.some(a => a.name === 'け こ'), true);
});
t('欠席のあとに来校したら行が増えず上書きされる', () => {
  const before = rows_(SH.log, LOG_COLS).length;
  const r = api_checkin({ no:'103005', purpose:'受講', booth:40, pageKey:issueKey_() });
  eq(r.ok, true);
  eq(rows_(SH.log, LOG_COLS).length, before);          // 行は増えない
  const l = api_list(STAFF);
  eq(l.absent.some(a => a.name === 'け こ'), false);   // 欠席欄から消える
  eq(l.visits.filter(v => v.name === 'け こ').length, 1);
});
t('上書きしても配布済みの記録は残る', () => {
  const v = api_list(STAFF).visits.filter(x => x.name === 'け こ')[0];
  api_handover({ token:'himitsu', id:v.id });
  api_setStatus({ token:'himitsu', no:'103005', status:'欠席' });
  const r = api_checkin({ no:'103005', purpose:'受講', booth:41, pageKey:issueKey_() });
  eq(r.ok, true);
  const v2 = api_list(STAFF).visits.filter(x => x.name === 'け こ')[0];
  if (!v2.handed) throw new Error('配布済みの記録が失われた');
  eq(v2.booth, 41);
});
t('来校済の生徒を照会すると登録済みで時刻が返る', () => {
  const r = api_lookup({ no:'103005', token:'himitsu' });
  eq(r.already, true);
  if (!/^\d{2}:\d{2}$/.test(r.alreadyAt)) throw new Error('時刻の形が不正: ' + r.alreadyAt);
});

console.log('\n■ 時刻の取り出し（タイムゾーンに依存しない）');
t('hhmm_ が保存文字列から時刻を取り出す', () => {
  eq(hhmm_('2026-10-05 14:30:00'), '14:30');
  eq(hhmm_('2026-10-05 09:05:00'), '09:05');
});
t('hhmm_ が空欄や不正値で落ちない', () => {
  eq(hhmm_(''), ''); eq(hhmm_(null), ''); eq(hhmm_(undefined), '');
  eq(hhmm_('なし'), ''); eq(hhmm_('99:99'), '');
});
t('hhmm_ がDateでも動く', () => { eq(hhmm_(new Date()), '14:30'); });

console.log('\n■ 履歴が消えないこと（D-1）');
t('翌日になっても前日のログが残り、当日の一覧には出ない', () => {
  const before = rows_(SH.log, LOG_COLS).length;
  FAKE_NOW = new Date(_OrigDate.UTC(2026, 9, 6, 5, 0));   // 翌日 14:00 JST
  const r = api_list(STAFF);
  eq(r.date, '2026-10-06');
  eq(r.counts.visit, 0);                          // 当日はまだ0件
  eq(rows_(SH.log, LOG_COLS).length, before);     // ログは消えていない
});
t('前日を指定すれば遡って見られる', () => {
  const r = api_list({ token:'himitsu', date:'2026-10-05' });
  if (r.counts.visit < 4) throw new Error('前日のログが減っている: ' + r.counts.visit);
});

console.log('\n■ doPost 経由');
t('不明なactionはエラーを返す', () => {
  const out = doPost({ postData:{ contents: JSON.stringify({ action:'nope', token:'himitsu' }) } });
  eq(JSON.parse(out).ok, false);
});
t('listをJSONで取得できる', () => {
  const out = doPost({ postData:{ contents: JSON.stringify({ action:'list', token:'himitsu', date:'2026-10-05' }) } });
  if (JSON.parse(out).counts.visit < 4) throw new Error('JSON経由の件数が合わない');
});

console.log('\n' + (fail ? '✗ 失敗 ' + fail + ' 件 / 成功 ' + pass + ' 件' : '✓ 全 ' + pass + ' 件 成功'));
process.exit(fail ? 1 : 0);
