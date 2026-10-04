/*
 * 来校管理システム 表の取り込みのテスト
 *
 *   node test/import.test.js
 *
 * gas/staff.html から読み取り用の関数を抜き出して評価する。
 * 実装を二重に持たないので、本体を直せばテストも追従する。
 *
 * 氏名はすべて架空のもの。実在の生徒のデータは含まない。
 */
const fs = require('fs');
const html = fs.readFileSync(process.argv[2] || __dirname + '/../gas/staff.html', 'utf8');

function grab(name){
  const i = html.indexOf('function ' + name + '(');
  if (i < 0) throw new Error('見つかりません: ' + name);
  let depth = 0, j = html.indexOf('{', i);
  for (let k = j; k < html.length; k++){
    if (html[k] === '{') depth++;
    else if (html[k] === '}'){ depth--; if (depth === 0) return html.slice(i, k + 1); }
  }
}
function grabVar(name){
  const i = html.search(new RegExp('var ' + name + '\\s*='));
  if (i < 0) throw new Error('見つかりません: ' + name);
  let depth = 0;
  for (let k = html.indexOf('{', i); k < html.length; k++){
    if (html[k] === '{') depth++;
    else if (html[k] === '}'){ depth--; if (depth === 0) return html.slice(i, k + 2); }
  }
}
// forEach の中で eval すると関数がその場のスコープに閉じるので、1つにまとめて一度に評価する
const SRC = [
  ['xlDate','xlTime','parseTable','stripHeader','width','pickCol','pickTextCol','pickNameCol',
   'colName','pickIdCol','detectRoster','detectRes'].map(grab).join('\n'),
  ['isNo','isGrade','isKana','isDate','isTime','isSchool'].map(grabVar).join('\n'),
  html.match(/var HEADER_WORDS = [^;]+;/)[0]
].join('\n');
eval(SRC);

let pass = 0, fail = 0;
const t = (n, f) => { try { f(); console.log('  ✓ ' + n); pass++; }
                      catch(e){ console.log('  ✗ ' + n + '\n      → ' + e.message); fail++; } };
const eq = (a, b, m) => { if (JSON.stringify(a) !== JSON.stringify(b))
  throw new Error((m || '') + ' 期待 ' + JSON.stringify(b) + ' / 実際 ' + JSON.stringify(a)); };

/* --- 本部システムの画面をコピーしたときの想定データ（架空） --- */
const RES_HEAD = ['校舎','番号','氏名','学年','生徒行事','在籍状況','予約講座','講数','予約日','予約時刻','ブース'];
const RES_BODY = [
  ['新松戸校','103101','安西　海斗','高３','マ','在籍','英語読解','2','2026/10/04','9:00','19'],
  ['新松戸校','103102','井上　garden'.replace('garden','結衣'),'高２','マ','在籍','数学IA','1','2026/10/04','9:00','56'],
  ['新松戸校','103103','上田　蓮','高３','マ','在籍','現代文','習1','2026/10/04','10:00','32'],
  ['新松戸校','103104','江口　美空','高１','マ','在籍','物理基礎','3','2026/10/04','17:30','15'],
  ['新松戸校','103105','大野　陽翔','高３','マ','在籍','古文','1','2026/10/04','19:00','5']
];
const ROS_HEAD = ['校舎コード','校舎名','マナビス生番号','氏名','氏名カナ','入会日','学校名','学年',
                  '生年月日','性別','保護者氏名','保護者氏名カナ','郵便番号','住所','電話番号'];
const ROS_BODY = [
  ['12345','新松戸校','103101','安西　海斗','アンザイ　カイト','2025/04/01','流山高校','高３','2008/07/15','男','安西　隆','アンザイ　タカシ','270-0000','千葉県','047-000-0001'],
  ['12345','新松戸校','103102','井上　結衣','イノウエ　ユイ','2025/05/13','松戸高校','高２','2009/07/10','女','井上　聡','イノウエ　サトシ','270-0001','千葉県','047-000-0002'],
  ['12345','新松戸校','103103','上田　蓮','ウエダ　レン','2025/05/24','柏高校','高３','2008/07/07','男','上田　誠','ウエダ　マコト','270-0002','千葉県','047-000-0003'],
  ['12345','新松戸校','103104','江口　美空','エグチ　ミク','2025/08/10','流山高校','高１','2010/09/07','女','江口　仁','エグチ　ジン','270-0003','千葉県','047-000-0004'],
  ['12345','新松戸校','103105','大野　陽翔','オオノ　ハルト','2025/08/10','松戸高校','高３','2008/08/04','男','大野　修','オオノ　オサム','270-0004','千葉県','047-000-0005']
];
const tsv = rows => rows.map(r => r.join('\t')).join('\r\n');

console.log('\n■ 貼り付けの読み取り');
t('タブ区切りを読める', () => {
  const r = parseTable(tsv([RES_HEAD].concat(RES_BODY)));
  eq(r.length, 6); eq(r[0][2], '氏名');
});
t('カンマ区切りを読める', () => {
  const r = parseTable([RES_HEAD, RES_BODY[0]].map(x => x.join(',')).join('\n'));
  eq(r.length, 2); eq(r[1][1], '103101');
});
t('空行や前後の空白を落とす', () => {
  const r = parseTable('\n\n' + tsv(RES_BODY) + '\n\n');
  eq(r.length, 5);
});
t('表でない文字列は空を返す', () => { eq(parseTable('ただの文章です'), []); eq(parseTable(''), []); });
t('見出し行を検出して外す', () => {
  const h = stripHeader(parseTable(tsv([RES_HEAD].concat(RES_BODY))));
  eq(h.head[2], '氏名'); eq(h.body.length, 5);
});
t('見出しが無ければ全行を本体として扱う', () => {
  const h = stripHeader(parseTable(tsv(RES_BODY)));
  eq(h.head, null); eq(h.body.length, 5);
});

console.log('\n■ 来校予約の列の判定');
t('見出しありで正しく割り当てる', () => {
  const r = detectRes(parseTable(tsv([RES_HEAD].concat(RES_BODY))), '');
  eq(r.items.length, 5);
  eq(r.items[0], { no:'103101', date:'2026-10-04', name:'安西　海斗', grade:'高３', at:'09:00' });
});
t('見出し無しでも中身から割り当てる', () => {
  const r = detectRes(parseTable(tsv(RES_BODY)), '');
  eq(r.items.length, 5);
  eq(r.items[0], { no:'103101', date:'2026-10-04', name:'安西　海斗', grade:'高３', at:'09:00' });
});
t('校舎名（全員同じ値）を氏名と取り違えない', () => {
  const r = detectRes(parseTable(tsv(RES_BODY)), '');
  if (r.items.some(x => x.name === '新松戸校')) throw new Error('校舎名を氏名として拾った');
});
t('予約講座を氏名と取り違えない', () => {
  const r = detectRes(parseTable(tsv(RES_BODY)), '');
  if (r.items.some(x => /英語|数学|現代文|物理|古文/.test(x.name))) throw new Error('講座名を氏名として拾った');
});
t('日付の列が無ければ指定した日付を使う', () => {
  const noDate = RES_BODY.map(r => r.slice(0, 8).concat(r.slice(9)));
  const r = detectRes(parseTable(tsv(noDate)), '2026-10-07');
  eq(r.items.length, 5);
  eq(r.items.every(x => x.date === '2026-10-07'), true);
});
t('日付の列も指定日も無ければ取り込まない', () => {
  const noDate = RES_BODY.map(r => r.slice(0, 8).concat(r.slice(9)));
  eq(detectRes(parseTable(tsv(noDate)), '').items.length, 0);
});
t('年のない「10/4」形式も読める', () => {
  const y = new Date().getFullYear();
  const short = RES_BODY.map(r => { const c = r.slice(); c[8] = '10/4'; return c; });
  const r = detectRes(parseTable(tsv(short)), '');
  eq(r.items[0].date, y + '-10-04');
});
t('生徒番号が無ければエラーを知らせる', () => {
  const r = detectRes(parseTable(tsv(RES_BODY.map(x => [x[0], x[2], x[3]]))), '2026-10-04');
  if (!/生徒番号らしい列/.test(r.note)) throw new Error('案内が出ていない: ' + r.note);
});

console.log('\n■ 生徒名簿の列の判定');
t('見出しありで5項目を割り当てる', () => {
  const r = detectRoster(parseTable(tsv([ROS_HEAD].concat(ROS_BODY))));
  eq(r.items.length, 5);
  eq(r.items[0], { no:'103101', name:'安西　海斗', kana:'アンザイ　カイト', grade:'高３', school:'流山高校' });
});
t('見出し無しでも5項目を割り当てる', () => {
  const r = detectRoster(parseTable(tsv(ROS_BODY)));
  eq(r.items.length, 5);
  eq(r.items[0], { no:'103101', name:'安西　海斗', kana:'アンザイ　カイト', grade:'高３', school:'流山高校' });
});
t('★保護者氏名を生徒の氏名と取り違えない（見出し無し）', () => {
  const r = detectRoster(parseTable(tsv(ROS_BODY)));
  const guardians = ROS_BODY.map(x => x[10]);
  r.items.forEach(it => {
    if (guardians.indexOf(it.name) >= 0) throw new Error('保護者氏名を生徒氏名として拾った: ' + it.name);
  });
});
t('保護者カナを氏名カナと取り違えない', () => {
  const r = detectRoster(parseTable(tsv(ROS_BODY)));
  const gk = ROS_BODY.map(x => x[11]);
  r.items.forEach(it => {
    if (gk.indexOf(it.kana) >= 0) throw new Error('保護者カナを拾った: ' + it.kana);
  });
});
t('生年月日を入会日と取り違えても氏名には影響しない', () => {
  const r = detectRoster(parseTable(tsv(ROS_BODY)));
  eq(r.items[0].name, '安西　海斗');
});
t('★校舎コード（全員同じ5桁）を生徒番号と取り違えない', () => {
  const r = detectRoster(parseTable(tsv(ROS_BODY)));
  if (r.items.some(x => x.no === '12345')) throw new Error('校舎コードを生徒番号として拾った');
});
t('電話番号や郵便番号を生徒番号と取り違えない', () => {
  const r = detectRoster(parseTable(tsv(ROS_BODY)));
  eq(r.items.map(x => x.no), ['103101','103102','103103','103104','103105']);
});
t('必要な5項目しか取り出さない（住所や電話は入らない）', () => {
  const r = detectRoster(parseTable(tsv([ROS_HEAD].concat(ROS_BODY))));
  eq(Object.keys(r.items[0]).sort(), ['grade','kana','name','no','school']);
  const json = JSON.stringify(r.items);
  ['千葉県','047-000','270-0000','2008/07/15'].forEach(bad => {
    if (json.indexOf(bad) >= 0) throw new Error('余計な項目が混ざっている: ' + bad);
  });
});

console.log('\n■ 日付と時刻の変換');
t('Excelのシリアル値を変換する', () => { eq(xlDate(46299), '2026-10-04'); eq(xlTime(0.375), '09:00'); });
t('文字列を変換する', () => {
  eq(xlDate('2026/10/4'), '2026-10-04'); eq(xlDate('2026-10-04T00:00:00.000Z'), '2026-10-04');
  eq(xlTime('9:00:00'), '09:00'); eq(xlTime('18:30'), '18:30');
});
t('1900年台のシリアル値は弾く', () => { eq(xlDate(60), ''); eq(xlDate(61), ''); });
t('日付として考えにくい値は弾く', () => {
  eq(xlDate(5), ''); eq(xlDate(2958465), ''); eq(xlDate(0), ''); eq(xlDate(NaN), '');
  eq(xlDate('2026/13/40'), '');
});
t('時刻として不正な値は弾く', () => { eq(xlTime('25:00'), ''); eq(xlTime('12:70'), ''); eq(xlTime(-1), ''); });
t('空値で壊れない', () => { eq(xlDate(''), ''); eq(xlTime(''), ''); eq(xlDate(null), ''); eq(xlTime(undefined), ''); });

console.log('\n' + (fail ? '✗ 失敗 ' + fail + ' 件 / 成功 ' + pass + ' 件' : '✓ 全 ' + pass + ' 件 成功'));
process.exit(fail ? 1 : 0);
