const fs=require('fs');
/* 本部システムのエクスポートの見出し行。個人情報は含まない（列名と日付のみ）。 */
const H = {
  "生徒名簿貼り付け": [
    "校舎コード",
    "校舎名",
    "マナビス生番号",
    "氏名",
    "氏名カナ",
    "入会日",
    "学校名",
    "学年",
    "生年月日",
    "性別",
    "保護者氏名",
    "保護者氏名カナ",
    "郵便番号",
    "住所",
    "電話番号",
    "担当アドバイザー",
    "生徒区分",
    "在籍状況",
    ""
  ],
  "来校予約貼り付け": [
    "校舎",
    "番号",
    "氏名",
    "学年",
    "生徒行事",
    "在籍状況",
    "予約講座",
    "講数",
    "予約日",
    "予約時刻",
    "ブース",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    ""
  ],
  "未予約貼り付け": [
    "校舎",
    "番号",
    "氏名",
    "学年",
    "生徒行事",
    "在籍状況",
    "予約講座",
    "講数",
    "予約日",
    "予約時刻",
    "ブース",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    ""
  ],
  "_serial_予約日": 46299,
  "_serial_予約時刻": 0.375,
  "_expect_予約日": "2026-10-04",
  "_expect_予約時刻": "09:00"
};

// staff.html から取り込み用の関数だけ抜き出して評価する（実装の二重管理を避ける）
const html=fs.readFileSync(process.argv[2] || __dirname+'/../gas/staff.html','utf8');
function grab(name){
  const i=html.indexOf('function '+name+'(');
  if(i<0) throw new Error('見つかりません: '+name);
  let depth=0, j=html.indexOf('{', i);
  for(let k=j;k<html.length;k++){
    if(html[k]==='{')depth++;
    else if(html[k]==='}'){depth--; if(depth===0) return html.slice(i,k+1);}
  }
}
eval(grab('xlDate')); eval(grab('xlTime')); eval(grab('mapCols'));

let pass=0,fail=0;
const t=(n,f)=>{try{f();console.log('  ✓ '+n);pass++}catch(e){console.log('  ✗ '+n+'\n      → '+e.message);fail++}};
const eq=(a,b,m)=>{if(JSON.stringify(a)!==JSON.stringify(b))throw new Error((m||'')+' 期待 '+JSON.stringify(b)+' / 実際 '+JSON.stringify(a))};

console.log('\n■ 実ファイルの見出しに対する列マッピング');
t('生徒名簿の5項目を正しく拾う', ()=>{
  const m=mapCols([H['生徒名簿貼り付け']], {
    no:['マナビス生番号','生徒番号','番号','生番号'], name:['氏名','生徒氏名','名前'],
    kana:['氏名カナ','フリガナ','ふりがな','カナ'], grade:['学年'], school:['学校名','高校名','在籍校']
  });
  if(!m) throw new Error('見出し行を認識できなかった');
  const h=H['生徒名簿貼り付け'];
  eq(h[m.idx.no],'マナビス生番号'); eq(h[m.idx.name],'氏名');
  eq(h[m.idx.kana],'氏名カナ'); eq(h[m.idx.grade],'学年'); eq(h[m.idx.school],'学校名');
});
t('氏名カナを氏名と誤認しない', ()=>{
  const h=H['生徒名簿貼り付け'];
  const m=mapCols([h], { no:['マナビス生番号'], name:['氏名'], kana:['氏名カナ'], grade:['学年'], school:['学校名'] });
  if(m.idx.name===m.idx.kana) throw new Error('氏名と氏名カナが同じ列になった');
  eq(h[m.idx.name],'氏名');
});
t('保護者氏名を氏名と誤認しない', ()=>{
  const h=H['生徒名簿貼り付け'];
  const m=mapCols([h], { no:['マナビス生番号'], name:['氏名'], kana:['氏名カナ'], grade:['学年'], school:['学校名'] });
  if(h[m.idx.name]==='保護者氏名') throw new Error('保護者氏名を拾った');
});
t('来校予約の5項目を正しく拾う', ()=>{
  const h=H['来校予約貼り付け'];
  const m=mapCols([h], {
    no:['番号','マナビス生番号','生徒番号','生番号'], name:['氏名','生徒氏名','名前'],
    grade:['学年'], date:['予約日','来校日','日付'], at:['予約時刻','来校予定時刻','時刻','時間']
  });
  if(!m) throw new Error('見出し行を認識できなかった');
  eq(h[m.idx.no],'番号'); eq(h[m.idx.name],'氏名'); eq(h[m.idx.grade],'学年');
  eq(h[m.idx.date],'予約日'); eq(h[m.idx.at],'予約時刻');
});
t('見出しが2行目以降にあっても探せる', ()=>{
  const m=mapCols([[],[],H['来校予約貼り付け']], {
    no:['番号'], name:['氏名'], grade:['学年'], date:['予約日'], at:['予約時刻']
  });
  eq(m.headRow, 2);
});
t('見出しが無ければ null を返す', ()=>{
  eq(mapCols([['あ','い','う']], { no:['番号'], name:['氏名'], grade:['学年'], date:['予約日'], at:['予約時刻'] }), null);
});

console.log('\n■ Excelシリアル値の変換（実データの値で検証）');
t('予約日のシリアル値が正しい日付になる', ()=>{ eq(xlDate(H['_serial_予約日']), H['_expect_予約日']); });
t('予約時刻のシリアル値が正しい時刻になる', ()=>{ eq(xlTime(H['_serial_予約時刻']), H['_expect_予約時刻']); });
t('文字列の日付も読める', ()=>{ eq(xlDate('2026/10/4'),'2026-10-04'); eq(xlDate('2026-10-04T00:00:00.000Z'),'2026-10-04'); });
t('文字列の時刻も読める', ()=>{ eq(xlTime('9:00:00'),'09:00'); eq(xlTime('18:30'),'18:30'); });
t('空値で壊れない', ()=>{ eq(xlDate(''),''); eq(xlTime(''),''); eq(xlDate(null),''); eq(xlTime(undefined),''); });
t('1900年台のシリアル値は一律で弾く（うるう年のずれを持ち込まない）', ()=>{
  eq(xlDate(60),'');                  // 存在しない1900-02-29
  eq(xlDate(61),'');                  // 1900-03-01
  eq(xlDate(59),'');
});
t('日付として考えにくい値は取り込まない', ()=>{
  eq(xlDate(5),'');                   // 1900-01-04 相当。迷子の数値セル
  eq(xlDate(2958465),'');             // 9999年
  eq(xlDate(0),''); eq(xlDate(-1),''); eq(xlDate(NaN),'');
  eq(xlDate('2026/13/40'),'');        // ありえない月日
});
t('時刻として不正な値は取り込まない', ()=>{
  eq(xlTime('25:00'),''); eq(xlTime('12:70'),''); eq(xlTime(-1),'');
});
t('実データのシリアル値と文字列が一致する', ()=>{
  eq(xlDate(H['_serial_予約日']), H['_expect_予約日']);
  eq(xlDate('2026/10/4'),'2026-10-04');
  eq(xlDate(H['_serial_予約日']), xlDate('2026/10/4'));
});

console.log('\n'+(fail?'✗ 失敗 '+fail+' 件 / 成功 '+pass+' 件':'✓ 全 '+pass+' 件 成功'));
process.exit(fail?1:0);
