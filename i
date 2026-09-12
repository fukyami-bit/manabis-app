<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<meta name="apple-mobile-web-app-capable" content="yes">
<title>マナビス新松戸校 業務アプリ</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"></script>
<style>

  html,body{margin:0;padding:0;background:#EFEDE7}
  #modebar{display:flex;gap:0;background:#12212E;position:sticky;top:0;z-index:900}
  #modebar button{flex:1;background:none;border:none;color:#9BA7B4;font-size:13px;font-weight:600;
    padding:11px 8px;cursor:pointer;letter-spacing:.03em;border-bottom:3px solid transparent;
    font-family:-apple-system,"Hiragino Kaku Gothic ProN","Yu Gothic Medium","Noto Sans JP",sans-serif}
  #modebar button.on{color:#fff;border-bottom-color:#3FBFA9;background:rgba(255,255,255,.06)}
  #modebar button small{display:block;font-weight:400;font-size:10.5px;opacity:.7;margin-top:1px}
  #appA,#appB{display:none}
  #appA.on,#appB.on{display:block}
  @media print{#modebar{display:none}}

/* ===== 面談スケジュール表・今週のAA ===== */
#appA{
    --paper:#F4F2EC; --sheet:#FFFFFF; --ink:#141C26; --ink-2:#4A5566; --line:#D3CEC2;
    --band-off:#D9D9D9; --band-on:#FFFFFF; --cyan:#CCFFFF;
    --teal:#0E6E6B; --teal-dark:#0A514F; --warn:#A24A12; --warn-bg:#FBEEE2;
    --mono:ui-monospace,"SFMono-Regular",Menlo,Consolas,monospace;
    --gothic:"Hiragino Kaku Gothic ProN","Yu Gothic Medium","Yu Gothic","Noto Sans JP",sans-serif;
    --mincho:"Hiragino Mincho ProN","Yu Mincho","YuMincho","Noto Serif JP",serif;
  }#appA *{box-sizing:border-box}#appA, #appA{margin:0;padding:0}#appA{background:var(--paper);color:var(--ink);font-family:var(--gothic);font-size:14px;line-height:1.7;-webkit-font-smoothing:antialiased}#appA .wrap{max-width:1180px;margin:0 auto;padding:28px 20px 80px}#appA header.masthead{display:flex;align-items:flex-end;justify-content:space-between;gap:20px;border-bottom:2px solid var(--ink);padding-bottom:12px;margin-bottom:26px}#appA h1{font-family:var(--mincho);font-size:26px;font-weight:600;letter-spacing:.06em;margin:0}#appA .masthead p{margin:4px 0 0;color:var(--ink-2);font-size:12.5px;letter-spacing:.02em}#appA .ruler{display:flex;gap:2px;align-items:flex-end;height:26px;flex-shrink:0}#appA .ruler i{display:block;width:6px;background:var(--band-off);border-radius:1px}#appA .ruler i.on{background:var(--teal)}#appA .cols{display:grid;grid-template-columns:minmax(0,380px) minmax(0,1fr);gap:26px;align-items:start}
  @media (max-width:900px){#appA .cols{grid-template-columns:1fr}}#appA .card{background:var(--sheet);border:1px solid var(--line);border-radius:2px;padding:18px 18px 20px;margin-bottom:18px}#appA .step{display:flex;align-items:baseline;gap:10px;margin:0 0 12px;font-size:14px;font-weight:600;letter-spacing:.04em}#appA .step span{font-family:var(--mono);font-size:11px;color:var(--teal);border:1px solid var(--teal);border-radius:999px;padding:1px 7px;font-weight:600}#appA .hint{color:var(--ink-2);font-size:12.5px;margin:0 0 12px}#appA .drop{border:1.5px dashed var(--line);border-radius:2px;background:#FCFBF8;padding:20px 16px;text-align:center;cursor:pointer;transition:border-color .15s,background .15s}#appA .drop:hover, #appA .drop.hot{border-color:var(--teal);background:#F2F8F7}#appA .drop strong{display:block;font-size:13.5px}#appA .drop em{display:block;font-style:normal;color:var(--ink-2);font-size:12px;margin-top:4px}#appA .drop:focus-visible{outline:2px solid var(--teal);outline-offset:2px}#appA .filed{display:flex;align-items:center;justify-content:space-between;gap:12px;background:#F2F8F7;border:1px solid #BFD8D6;border-radius:2px;padding:10px 12px}#appA .filed .nm{font-size:13px;overflow:hidden}#appA .filed .nm b{font-weight:600;display:block;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}#appA .filed .nm small{display:block;color:var(--ink-2);font-size:11.5px}#appA button{font-family:inherit;font-size:13px;cursor:pointer}#appA .btn{background:var(--teal);color:#fff;border:none;border-radius:2px;padding:11px 18px;font-weight:600;letter-spacing:.04em;transition:background .15s}#appA .btn:hover:not(:disabled){background:var(--teal-dark)}#appA .btn:disabled{background:#B9BEC2;cursor:not-allowed}#appA .btn.wide{width:100%;padding:14px;font-size:14.5px}#appA .btn.alt{margin-top:8px;background:#fff;color:#C4463C;border:1.5px solid #F09189}#appA .btn.alt:hover:not(:disabled){background:#FDF1EF}#appA .btn.alt:disabled{background:#F5F4F0;color:#B9BEC2;border-color:#DDD9CF}#appA .btn-ghost{background:none;border:1px solid var(--line);border-radius:2px;padding:6px 11px;color:var(--ink-2);white-space:nowrap}#appA .btn-ghost:hover{border-color:var(--ink-2);color:var(--ink)}#appA label.field{display:block;margin-bottom:10px}#appA label.field span{display:block;font-size:12px;color:var(--ink-2);margin-bottom:3px}#appA input[type=text], #appA input[type=number], #appA select{width:100%;font-family:inherit;font-size:13px;padding:7px 9px;border:1px solid var(--line);border-radius:2px;background:#fff;color:var(--ink)}#appA input:focus, #appA select:focus{outline:2px solid var(--teal);outline-offset:-1px;border-color:var(--teal)}#appA .row2{display:grid;grid-template-columns:1fr 1fr;gap:10px}#appA .pattern{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:8px 10px;margin-top:6px}#appA .pattern label span{font-size:11.5px;color:var(--ink-2)}#appA .pattern input{font-family:var(--mono);font-size:12px;padding:5px 7px}#appA table.staff{width:100%;border-collapse:collapse;font-size:12.5px}#appA table.staff th{text-align:left;font-weight:600;color:var(--ink-2);font-size:11.5px;border-bottom:1px solid var(--line);padding:5px 6px;white-space:nowrap;background:var(--sheet);position:sticky;top:0}#appA table.staff td{border-bottom:1px solid #EDEAE2;padding:3px 6px}#appA table.staff td.d{font-family:var(--mono);white-space:nowrap;font-size:12px}#appA table.staff td.d.we{color:var(--warn)}#appA table.staff input{font-family:var(--mono);font-size:12px;padding:4px 6px;min-width:96px}#appA table.staff th button{background:none;border:none;color:var(--ink-2);padding:0 0 0 4px;font-size:12px}#appA table.staff th button:hover{color:var(--warn)}#appA .preview-head{display:flex;align-items:baseline;justify-content:space-between;gap:12px;margin-bottom:10px}#appA .preview-head h2{font-family:var(--mincho);font-size:17px;margin:0;letter-spacing:.05em}#appA .scale{display:grid;grid-template-columns:74px minmax(0,1fr);gap:10px;font-family:var(--mono);font-size:10px;color:var(--ink-2);margin-bottom:2px}#appA .scale .ticks{display:grid;grid-template-columns:repeat(26,1fr)}#appA .scale .ticks i{grid-column:span 2;font-style:normal;border-left:1px solid var(--line);padding-left:3px}#appA .day{display:grid;grid-template-columns:74px minmax(0,1fr);gap:10px;padding:7px 0;border-top:1px solid #E7E3DA;align-items:start}#appA .day.alert{background:var(--warn-bg);margin:0 -8px;padding-left:8px;padding-right:8px;border-radius:2px}#appA .daylabel{font-family:var(--mono);font-size:12px;padding-top:2px}#appA .daylabel b{font-weight:600}#appA .daylabel small{display:block;color:var(--ink-2);font-size:10.5px;font-family:var(--gothic)}#appA .rows{display:flex;flex-direction:column;gap:3px}#appA .row{display:grid;grid-template-columns:64px minmax(0,1fr);gap:6px;align-items:center}#appA .row .who{font-size:11.5px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;text-align:right}#appA .row .who.staff{font-weight:600}#appA .row .who.off{color:#A6ADB6}#appA .bands{display:grid;grid-template-columns:repeat(26,1fr);height:14px;border:1px solid #CFCABE}#appA .bands i{display:block;background:var(--band-off);border-right:1px solid rgba(0,0,0,.06)}#appA .bands i.on{background:var(--band-on)}#appA .bands i:last-child{border-right:none}#appA .note{font-size:11.5px;color:var(--warn);margin-top:3px}#appA .empty{border:1px dashed var(--line);border-radius:2px;padding:44px 20px;text-align:center;color:var(--ink-2);background:repeating-linear-gradient(180deg,#fff 0 22px,#FBFAF7 22px 44px)}#appA .empty b{display:block;color:var(--ink);font-family:var(--mincho);font-size:15px;margin-bottom:6px;letter-spacing:.04em}#appA .alerts{margin:0 0 14px;padding:0;list-style:none}#appA .alerts li{background:var(--warn-bg);border-left:3px solid var(--warn);padding:7px 11px;font-size:12.5px;margin-bottom:5px}#appA .status{font-size:12.5px;color:var(--ink-2);margin-top:10px;min-height:1.6em}#appA .status.err{color:var(--warn)}#appA .legend{display:flex;gap:14px;font-size:11.5px;color:var(--ink-2);margin-top:12px;flex-wrap:wrap}#appA .legend i{display:inline-block;width:14px;height:11px;border:1px solid #CFCABE;vertical-align:-1px;margin-right:4px}#appA .legend i.off{background:var(--band-off)}#appA .legend i.on{background:#fff}#appA details.help, #appA details.check{margin-top:14px;font-size:12.5px;color:var(--ink-2)}#appA details summary{cursor:pointer;color:var(--teal)}#appA details code{font-family:var(--mono);background:#F0EEE7;padding:1px 4px;border-radius:2px}#appA details ul{padding-left:18px;margin:8px 0}#appA .scroll{max-height:360px;overflow:auto;border:1px solid var(--line);margin-top:10px}

/* ===== ターム面談の日程設定 ===== */
#appB{
    --ink:#12212E; --ink-2:#5A6672; --line:#DFDCD4; --paper:#F5F3EE; --card:#FFFFFF;
    --ok:#1B7A5A; --ok-bg:#E4F2EC; --warn:#B06A0F; --warn-bg:#FBF0DC; --ng:#B23A2E; --ng-bg:#FAE7E4;
    --accent:#0E6E6B; --accent-dark:#0A514F;
    --gothic:-apple-system,"Hiragino Kaku Gothic ProN","Yu Gothic Medium","Noto Sans JP",sans-serif;
    --num:ui-monospace,"SF Mono",Menlo,monospace;
    --tab-h:64px;
  }#appB *{box-sizing:border-box;-webkit-tap-highlight-color:transparent}#appB, #appB{margin:0;padding:0;background:var(--paper);color:var(--ink);font-family:var(--gothic);
    font-size:15px;line-height:1.6;-webkit-font-smoothing:antialiased}#appB{padding-bottom:calc(var(--tab-h) + env(safe-area-inset-bottom))}#appB button, #appB input, #appB select{font-family:inherit;font-size:inherit;color:inherit}#appB button{cursor:pointer;border:none;background:none}#appB header{position:sticky;top:0;z-index:20;background:var(--ink);color:#fff;
    padding:calc(env(safe-area-inset-top) + 10px) 16px 10px}#appB .hd{display:flex;align-items:baseline;justify-content:space-between;gap:10px}#appB .hd h1{margin:0;font-size:17px;font-weight:600;letter-spacing:.04em}#appB .hd .mon{font-family:var(--num);font-size:13px;opacity:.75}#appB .hd .badge{background:#E0A73C;color:#12212E;font-size:12px;font-weight:700;
    padding:2px 9px;border-radius:999px;font-family:var(--num)}#appB .stale{background:#7A3B2E;color:#fff;font-size:12.5px;padding:6px 16px;text-align:center}#appB main{padding:14px 14px 24px;max-width:640px;margin:0 auto}#appB .empty{background:var(--card);border:1px dashed var(--line);border-radius:10px;
    padding:34px 20px;text-align:center;color:var(--ink-2);font-size:13.5px}#appB .empty b{display:block;color:var(--ink);font-size:15px;margin-bottom:8px}#appB .search{display:flex;gap:8px;margin-bottom:10px}#appB .search input{flex:1;padding:12px 14px;border:1px solid var(--line);border-radius:10px;background:#fff}#appB .search input:focus{outline:2px solid var(--accent);outline-offset:-1px}#appB .chips{display:flex;gap:6px;overflow-x:auto;padding-bottom:10px;-webkit-overflow-scrolling:touch}#appB .chip{flex-shrink:0;padding:7px 13px;border-radius:999px;border:1px solid var(--line);
    background:#fff;font-size:13px;white-space:nowrap}#appB .chip.on{background:var(--ink);color:#fff;border-color:var(--ink)}#appB .stu{display:flex;align-items:center;gap:12px;width:100%;text-align:left;background:var(--card);
    border:1px solid var(--line);border-left:5px solid var(--tc,#CCC);border-radius:10px;
    padding:12px 14px;margin-bottom:8px}#appB .stu .nm{flex:1;min-width:0}#appB .stu .nm b{display:block;font-size:16px;font-weight:600}#appB .stu .nm small{color:var(--ink-2);font-size:12px;font-family:var(--num)}#appB .stu .rest{flex-shrink:0;text-align:center;min-width:52px}#appB .stu .rest b{display:block;font-family:var(--num);font-size:20px;line-height:1.1}#appB .stu .rest small{font-size:10.5px;color:var(--ink-2)}#appB .stu.done .rest b{color:var(--ok)}#appB .stu.need .rest b{color:var(--ng)}#appB .sheet{position:fixed;inset:0;z-index:40;background:rgba(18,33,46,.45);
    display:flex;align-items:flex-end;justify-content:center}#appB .sheet-in{background:var(--paper);width:100%;max-width:640px;max-height:92vh;
    border-radius:16px 16px 0 0;display:flex;flex-direction:column;overflow:hidden}#appB .sheet-hd{padding:14px 16px;background:var(--card);border-bottom:1px solid var(--line);
    display:flex;align-items:flex-start;gap:12px}#appB .sheet-hd .t{flex:1;min-width:0}#appB .sheet-hd .t b{font-size:18px}#appB .sheet-hd .t small{display:block;color:var(--ink-2);font-size:12.5px}#appB .sheet-hd .x{font-size:15px;color:var(--ink-2);padding:4px 10px;border:1px solid var(--line);border-radius:8px;background:#fff}#appB .sheet-bd{overflow-y:auto;padding:12px 14px calc(20px + env(safe-area-inset-bottom));-webkit-overflow-scrolling:touch}#appB .setdates{padding:10px 16px;background:#EFEDE6;font-size:13px;border-bottom:1px solid var(--line)}#appB .setdates span{font-family:var(--num);font-weight:600;margin-right:10px}#appB .day{display:flex;align-items:center;gap:12px;width:100%;text-align:left;background:var(--card);
    border:1px solid var(--line);border-radius:10px;padding:11px 13px;margin-bottom:7px}#appB .day.mine{border-color:var(--accent);background:#F1F8F7}#appB .day.full{opacity:.55}#appB .day .d{flex-shrink:0;width:62px;font-family:var(--num)}#appB .day .d b{font-size:21px;font-weight:600;line-height:1}#appB .day .d small{display:block;font-size:11px;color:var(--ink-2)}#appB .day .d.sat b, #appB .day .d.sat small{color:#3B57C4}#appB .day .d.sun b, #appB .day .d.sun small{color:#C0392B}#appB .day .info{flex:1;min-width:0}#appB .day .info .bar{height:7px;border-radius:4px;background:#E6E3DC;overflow:hidden;margin-bottom:5px}#appB .day .info .bar i{display:block;height:100%}#appB .day .info small{font-size:11.5px;color:var(--ink-2);display:block;
    white-space:nowrap;overflow:hidden;text-overflow:ellipsis}#appB .day .rest{flex-shrink:0;text-align:center;min-width:56px;padding:5px 4px;border-radius:8px;font-size:11px}#appB .day .rest b{display:block;font-family:var(--num);font-size:17px;line-height:1.1}#appB .r-ok{background:var(--ok-bg);color:var(--ok)}#appB .r-warn{background:var(--warn-bg);color:var(--warn)}#appB .r-ng{background:var(--ng-bg);color:var(--ng)}#appB .note{font-size:12px;color:var(--warn);margin:2px 0 8px 74px}#appB table.grid{width:100%;border-collapse:collapse;font-size:12.5px;background:var(--card)}#appB table.grid th{position:sticky;top:0;background:var(--ink);color:#fff;padding:7px 4px;font-weight:600;font-size:12px}#appB table.grid td{border-bottom:1px solid #EEEBE4;padding:5px 3px;text-align:center;font-family:var(--num)}#appB table.grid td.dt{text-align:left;padding-left:8px;white-space:nowrap}#appB table.grid td.sat{color:#3B57C4}#appB table.grid td.sun{color:#C0392B}#appB td .cell{display:block;border-radius:5px;padding:3px 0;font-size:12px}#appB .card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:16px;margin-bottom:12px}#appB .card h2{margin:0 0 4px;font-size:15px;font-weight:600}#appB .card p{margin:0 0 12px;font-size:12.5px;color:var(--ink-2)}#appB .file{display:block;width:100%;padding:14px;border:1.5px dashed var(--line);border-radius:10px;
    background:#FCFBF8;text-align:center;font-size:14px}#appB .file.set{border-style:solid;border-color:#BFD8D6;background:#F1F8F7}#appB .btn{display:block;width:100%;padding:14px;background:var(--accent);color:#fff;border-radius:10px;
    font-weight:600;font-size:15px;text-align:center}#appB .btn:disabled{background:#B9BEC2}#appB .btn.sub{background:#fff;color:var(--accent);border:1.5px solid #BFD8D6}#appB .btn.warnbtn{background:#fff;color:var(--ng);border:1.5px solid #E8C4BE}#appB .row2{display:flex;gap:8px}#appB .row2>*{flex:1}#appB label.f{display:block;margin-bottom:10px;font-size:12.5px;color:var(--ink-2)}#appB label.f input{width:100%;padding:10px 12px;border:1px solid var(--line);border-radius:9px;background:#fff;
    color:var(--ink);font-size:14px;margin-top:4px}#appB pre.out{background:#12212E;color:#E8E6E0;font-family:var(--num);font-size:12px;
    padding:12px;border-radius:9px;overflow-x:auto;white-space:pre;margin:0 0 10px}#appB .pend{display:flex;align-items:center;gap:10px;padding:9px 0;border-bottom:1px solid #EEEBE4;font-size:14px}#appB .pend .d{font-family:var(--num);font-weight:600;width:52px}#appB .pend .n{flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}#appB .pend button{color:var(--ng);font-size:12.5px;padding:4px 8px;border:1px solid #E8C4BE;border-radius:7px}#appB nav{position:fixed;left:0;right:0;bottom:0;z-index:30;background:rgba(255,255,255,.96);
    backdrop-filter:blur(10px);border-top:1px solid var(--line);
    display:flex;padding-bottom:env(safe-area-inset-bottom)}#appB nav button{flex:1;padding:9px 2px 8px;font-size:10.5px;color:var(--ink-2);display:flex;
    flex-direction:column;align-items:center;gap:3px}#appB nav button.on{color:var(--accent);font-weight:600}#appB nav svg{width:22px;height:22px;stroke:currentColor;fill:none;stroke-width:1.7;
    stroke-linecap:round;stroke-linejoin:round}#appB nav .dot{position:absolute;top:6px;margin-left:22px;width:7px;height:7px;border-radius:50%;background:#E0A73C}#appB .toast{position:fixed;left:50%;bottom:calc(var(--tab-h) + 20px);transform:translateX(-50%);
    background:var(--ink);color:#fff;padding:11px 18px;border-radius:999px;font-size:13.5px;
    z-index:60;opacity:0;transition:opacity .2s;pointer-events:none;white-space:nowrap;max-width:92vw}#appB .toast.show{opacity:1}

</style>
</head>
<body>
<div id="modebar">
  <button data-mode="B" id="mbB">ターム面談の日程設定<small>iPhone向け</small></button>
  <button data-mode="A" id="mbA">面談スケジュール表・今週のAA<small>PC向け</small></button>
</div>
<div id="appA">
<div class="wrap">
  <header class="masthead">
    <div>
      <h1>面談スケジュール表 自動作成</h1>
      <p>シフト表を読み込んで、1か月分のスケジュール表を書式そのままで書き出します。</p>
    </div>
    <div class="ruler" aria-hidden="true">
      <i style="height:8px"></i><i style="height:8px"></i><i style="height:14px" class="on"></i>
      <i style="height:14px" class="on"></i><i style="height:20px" class="on"></i><i style="height:26px" class="on"></i>
      <i style="height:20px" class="on"></i><i style="height:12px"></i>
    </div>
  </header>

  <div class="cols">
    <div>
      <div class="card">
        <p class="step"><span>1</span>テンプレート</p>
        <p class="hint">面談スケジュール表のExcel（<code>原版(触らない)</code>シートを含むもの）。一度読み込めばこの端末に保存され、次回からは不要です。</p>
        <div id="tplZone"></div>
      </div>

      <div class="card">
        <p class="step"><span>2</span>シフト表</p>
        <p class="hint">ExcelでもPDFでも読み込めます。日付が縦でも横でも、自動で向きを判別します。</p>
        <p class="hint" style="margin-top:-6px">スタッフ欄は5列固定です。職員2名が左端、アルバイトは勤務が重ならない人どうしを同じ列にまとめて3列に収めます。4名以上重なる日は、勤務終了が同じ人と相乗りさせます。</p>
        <div id="shiftZone"></div>
        <label class="field" id="monthField" style="display:none;margin-top:12px">
          <span>対象月</span>
          <select id="monthSel"></select>
        </label>
        <div id="ymFix" style="display:none;margin-top:10px">
          <p class="hint" style="margin-bottom:6px;color:var(--warn)">ファイルから年月を読み取れませんでした。正しい年月を指定してください。</p>
          <div class="row2">
            <label class="field"><span>年</span><input type="number" id="ymY" min="2000" max="2100"></label>
            <label class="field"><span>月</span><input type="number" id="ymM" min="1" max="12"></label>
          </div>
        </div>
        <details class="check" id="checkBox" style="display:none">
          <summary>読み取り結果を確認・修正する</summary>
          <p class="hint" style="margin-top:8px">セルを直接書き換えられます。見出しの × でスタッフ列を削除できます。</p>
          <div class="scroll"><table class="staff"><thead id="cellHead"></thead><tbody id="cellBody"></tbody></table></div>
        </details>
      </div>

      <div class="card" id="staffCard" style="display:none">
        <p class="step"><span>3</span>職員の勤務</p>
        <p class="hint">シフト表から読み取った内容が入ります。<b>数字だけの記入は勤務開始時刻</b>とみなし、そこから拘束時間ぶん（開館時間の内側）を出勤として塗ります。空欄・「休」「振休」などはその日休みです。</p>
        <div class="row2">
          <label class="field">
            <span>職員（左端に固定表示・カンマ区切り）</span>
            <input type="text" id="fixedNames" value="平野,小西">
          </label>
          <label class="field">
            <span>拘束時間（時間）</span>
            <input type="number" id="holdHours" value="9" min="1" max="14" step="0.5">
          </label>
        </div>
        <div style="margin-bottom:6px;font-size:12px;color:var(--ink-2)">曜日ごとの基本パターン（シフト表に職員の欄がないときに使います）</div>
        <div class="pattern" id="pattern"></div>
        <div style="display:flex;gap:8px;margin:10px 0 4px">
          <button class="btn-ghost" id="applyPattern">パターンを空欄に入れる</button>
          <button class="btn-ghost" id="clearStaff">全部消す</button>
        </div>
        <div class="scroll" style="max-height:340px"><table class="staff"><thead id="staffHead"></thead><tbody id="staffBody"></tbody></table></div>
      </div>

      <div class="card">
        <button class="btn wide" id="genBtn" disabled>面談スケジュール表を書き出す</button>
        <button class="btn wide alt" id="posterBtn" disabled>今週のAA（出勤予定表）を書き出す</button>
        <p class="status" id="status">テンプレートとシフト表を読み込むと書き出せます。</p>
        <details class="help">
          <summary>今週のAAの使い方</summary>
          <ul>
            <li>アルバイトのみを、同じ時間帯の人ごとにまとめて印刷用のページにします（平野・小西は載せません）</li>
            <li>書き出したファイルをブラウザで開き、印刷（Ctrl+P / ⌘P）で「送信先：PDFに保存」「用紙：A4」「余白：なし」「背景のグラフィック：オン」にすると見本と同じPDFになります</li>
            <li>印刷する前に、画面上で時間や名前を直接書き換えられます</li>
            <li>1ページ7日ぶん。日付の並び順のまま区切ります</li>
          </ul>
        </details>
        <details class="help">
          <summary>読み込めるシフト表について</summary>
          <ul>
            <li>Excel（.xlsx / .xlsm）とPDFに対応。PDFは文字情報を含むもの（Excelから書き出したPDFなど）に限ります</li>
            <li>日付が縦に並ぶ表でも、スタッフが縦・日付が横の表でも読み取れます</li>
            <li>時刻は <code>17:00-21:30</code> でも <code>18-21.75</code> でも可</li>
            <li>分割勤務は <code>9:00-13:00、18:00-22:00</code> のように区切る</li>
            <li><code>終日</code> は開館時間、空欄・<code>×</code>・<code>休</code> はその日勤務なし</li>
            <li>職員の欄が <code>12</code> のような数字だけの表でも、勤務開始時刻として読み取ります</li>
            <li>全角の数字やダッシュ（<code>９－１３</code>）、セル内の改行やメモ書きが混じっていても読み取れます</li>
            <li>紙をスキャンしただけのPDF（画像）は読み取れません</li>
          </ul>
        </details>
      </div>
    </div>

    <div>
      <div class="card" style="min-height:420px">
        <div class="preview-head">
          <h2 id="previewTitle">プレビュー</h2>
          <span style="font-size:12px;color:var(--ink-2)" id="previewCount"></span>
        </div>
        <ul class="alerts" id="alerts"></ul>
        <div id="preview">
          <div class="empty">
            <b>まだ読み込んでいません</b>
            シフト表を読み込むと、1日ずつの出勤帯をここで確認できます。<br>白い部分がスケジュール表の記入欄、グレーが塗りつぶしです。
          </div>
        </div>
        <div class="legend" id="legend" style="display:none">
          <span><i class="on"></i>出勤（記入欄）</span>
          <span><i class="off"></i>不在（グレー塗りつぶし）</span>
          <span>1目盛 = 30分（9:00〜22:00）</span>
          <span>同じ列の「A・B」は2人で1列を分け合う組み合わせです</span>
        </div>
      </div>
    </div>
  </div>
</div>


</div>
<div id="appB">

<header>
  <div class="hd">
    <h1>ターム面談 設定</h1>
    <span class="mon" id="hdMon">未読込</span>
    <span class="badge" id="hdBadge" style="display:none">0</span>
  </div>
</header>
<div class="stale" id="stale" style="display:none"></div>

<main id="view"></main>

<nav id="nav"></nav>
<div class="toast" id="toast"></div>



</div>
<script>
(function(){
  var KEY='manabis-app-mode';
  function apply(mode){
    document.getElementById('appA').className = (mode==='A'?'on':'');
    document.getElementById('appB').className = (mode==='B'?'on':'');
    document.getElementById('mbA').className = (mode==='A'?'on':'');
    document.getElementById('mbB').className = (mode==='B'?'on':'');
    try{ localStorage.setItem(KEY, mode); }catch(e){}
    window.scrollTo(0,0);
  }
  var saved=null;
  try{ saved=localStorage.getItem(KEY); }catch(e){}
  // 初回は画面の広さで判断（スマホなら日程設定、PCならスケジュール表）
  apply(saved || (window.innerWidth < 760 ? 'B' : 'A'));
  document.getElementById('mbA').addEventListener('click',function(){apply('A');});
  document.getElementById('mbB').addEventListener('click',function(){apply('B');});
})();
</script>
<script>
(function(){

// ===CORE START===
const FIRST_ROW = 3, LAST_ROW = 28, DAY_START = 9 * 60, SLOT = 30, SLOTS = 26;
const NAME_ROW = 2, MIN_PT = 3;
const DAY_END = 9 * 60 + 26 * 30;   // 表のいちばん下（22:00）
const SRC = { time: 1, narrow: 2, wide: 3, spacer: 12, b1: 13, b2: 14, tail: 15 };
const TPL_LAST_ROW = 33;
const GRAY = 'FFD9D9D9', WHITE = 'FFFFFFFF';
const SKIP_HEADERS = ['日付', '曜日', '日付/曜日', '日付・曜日', '開館時間', '日', '曜', '備考', '特記事項',
  '不足', '合計', 'メモ', 'note', '氏名', '名前', 'スタッフ', '文', '理', '時間', '必要人数'];
const STAFF_HOLD_HOURS = 9;   // 職員の拘束時間（勤務開始からこの時間ぶんが出勤）
const SKIP_HEADER_RE = /日付|曜日|特記|備考|不足|合計|メモ|人数|区分|氏名|名前|スタッフ|時間$/;
const WD = ['日', '月', '火', '水', '木', '金', '土'];
const WD_RE = /^[月火水木金土日]$/;

/* ---------- 小道具 ---------- */
function cellText(v) {
  if (v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return v.getUTCFullYear() + '-' + String(v.getUTCMonth() + 1).padStart(2, '0') + '-' + String(v.getUTCDate()).padStart(2, '0');
  }
  if (typeof v === 'object') {
    if (v.richText) return v.richText.map(t => t.text).join('');
    if (v.result !== undefined && v.result !== null) return cellText(v.result);
    if (v.text !== undefined) return String(v.text);
    return '';
  }
  return String(v);
}

/** 全角の数字・記号を半角にそろえる */
function normalizeText(v) {
  return String(v == null ? '' : v)
    .replace(/[０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0))
    .replace(/[－ー―‐–—〜～~]/g, '-')
    .replace(/[．]/g, '.').replace(/[：]/g, ':').replace(/[／]/g, '/')
    .replace(/[（]/g, '(').replace(/[）]/g, ')');
}

function toMinutes(t) {
  t = String(t).trim();
  if (t.indexOf(':') >= 0) { const p = t.split(':'); return parseInt(p[0], 10) * 60 + parseInt(p[1], 10); }
  return Math.round(parseFloat(t) * 60);
}

/** セル内で折り返された文字をつなぐ。時間帯が途中で切れている場合はそのまま連結する */
function joinCellText(a, b) {
  if (!a) return b || '';
  if (!b) return a;
  if (/-\s*$/.test(a) || /^\s*-/.test(b)) return a + b;
  return a + '、' + b;
}

/** セル文字列から勤務時間帯 [[開始分,終了分],...] */
function parseRanges(text, open, close) {
  if (text === null || text === undefined) return [];
  const s = normalizeText(text).trim();
  if (!s) return [];
  if (['×', 'x', 'X', '✕', '-', '−', '休', '休み', 'なし'].indexOf(s) >= 0) return [];
  if (s.indexOf('終日') >= 0 || s.indexOf('通し') >= 0) return (open == null) ? [] : [[open, close]];
  const out = [];
  s.split(/[\n,、／\/]+/).forEach(seg => {
    const m = seg.match(/(\d{1,2}(?::\d{2}|\.\d+)?)\s*-\s*(\d{1,2}(?::\d{2}|\.\d+)?)/);
    if (!m) return;
    const a = toMinutes(m[1]), b = toMinutes(m[2]);
    if (b > a && a >= 0 && b <= 24 * 60) out.push([a, b]);
  });
  return out.sort((p, q) => p[0] - q[0]);
}

/** 「12」「8.5」のように数字だけの記入を勤務開始時刻として読む */
function parseStartOnly(text) {
  const s = normalizeText(text).trim();
  const m = s.match(/^(\d{1,2}(?:\.\d+)?)$/);
  if (!m) return null;
  const v = Math.round(parseFloat(m[1]) * 60);
  return (v >= 0 && v < 24 * 60) ? v : null;
}

/** 開館時間からはみ出す部分を落とす */
function clipRanges(ranges, open, close) {
  if (open == null || close == null) return ranges;
  const out = [];
  ranges.forEach(r => {
    const a = Math.max(r[0], open), b = Math.min(r[1], close);
    if (b > a) out.push([a, b]);
  });
  return out;
}

/** セル1つを勤務時間帯に変換する。
 *  hours を渡すと、数字だけの記入を「開始時刻＋hours時間」として扱う（職員用）。 */
function parseCell(text, open, close, hours, clip) {
  let rg = parseRanges(text, open, close);
  if (!rg.length && hours) {
    const st = parseStartOnly(text);
    if (st != null) rg = [[st, st + Math.round(hours * 60)]];
  }
  return clip === false ? rg : clipRanges(rg, open, close);
}

function isWorking(ranges, idx) {
  const s = DAY_START + idx * SLOT, e = s + SLOT;
  return ranges.some(r => r[0] < e && r[1] > s);
}

/** 日付らしさの判定。'full'=年月日 / 'md'=月日 / 'day'=日のみ */
function readDate(text) {
  const t = normalizeText(text).trim();
  if (!t) return null;
  let m = t.match(/^(\d{4})\s*[\/\-年.]\s*(\d{1,2})\s*[\/\-月.]\s*(\d{1,2})/);
  if (m) return { mode: 'full', y: +m[1], m: +m[2], d: +m[3] };
  m = t.match(/^(\d{1,2})\s*[\/\-月.]\s*(\d{1,2})\s*日?\s*[(（]?/);
  if (m && +m[1] >= 1 && +m[1] <= 12 && +m[2] >= 1 && +m[2] <= 31) return { mode: 'md', m: +m[1], d: +m[2] };
  m = t.match(/^(\d{1,2})\s*日?\s*[(（]?[月火水木金土日]?[)）]?$/);
  if (m && +m[1] >= 1 && +m[1] <= 31) return { mode: 'day', d: +m[1] };
  return null;
}

/** 「2026年9月」「9月」などから年月を拾う */
function detectYearMonth(allText) {
  let m = allText.match(/(20\d{2})\s*年\s*(\d{1,2})\s*月/);
  if (m) return { y: +m[1], m: +m[2] };
  m = allText.match(/(20\d{2})[\/\-](\d{1,2})(?![\/\-]\d)/);
  if (m) return { y: +m[1], m: +m[2] };
  m = allText.match(/(\d{1,2})\s*月/);
  if (m) return { y: null, m: +m[1] };
  return null;
}

const transpose = g => {
  const w = g.reduce((a, r) => Math.max(a, r.length), 0);
  return Array.from({ length: w }, (_, c) => g.map(r => r[c] || ''));
};

/* ---------- Excel → グリッド ---------- */
function gridFromWorksheet(ws) {
  const rows = [];
  const maxC = Math.min(ws.columnCount || 40, 80);
  for (let r = 1; r <= ws.rowCount; r++) {
    const row = ws.getRow(r), out = [];
    for (let c = 1; c <= maxC; c++) out.push(cellText(row.getCell(c).value).trim());
    rows.push(out);
  }
  return rows;
}

/* ---------- PDF → グリッド ---------- */
/** 1ページ分のテキスト片を、行×列のグリッドに組み直す */
function gridFromItems(items) {
  const src = items.filter(it => String(it.s).trim() !== '');
  if (!src.length) return { grid: [], ys: [] };
  const hs = src.map(it => it.h).sort((a, b) => a - b);
  const H = hs[Math.floor(hs.length / 2)] || 10;

  // 行：y座標が近いものをまとめる
  const sorted = src.slice().sort((a, b) => (b.y - a.y) || (a.x - b.x));
  const lines = [];
  sorted.forEach(it => {
    const last = lines[lines.length - 1];
    if (last && Math.abs(last.y - it.y) <= H * 0.6) { last.items.push(it); last.y = (last.y + it.y) / 2; }
    else lines.push({ y: it.y, items: [it] });
  });

  // 列：セルの左端（または中央）がそろう位置を列の基準にする。
  //     セル内の文字が隣の空きセルにはみ出しているPDFでも列を取り違えないため。
  const rowsCount = lines.length;
  const clusterBy = pick => {
    const pts = [];
    lines.forEach((line, li) => line.items.forEach(it => pts.push({ v: pick(it), li: li })));
    pts.sort((a, b) => a.v - b.v);
    const cl = [];
    pts.forEach(pt => {
      const last = cl[cl.length - 1];
      if (last && pt.v - last.v <= H * 0.6) { last.v = (last.v * last.n + pt.v) / (last.n + 1); last.n++; last.rows.add(pt.li); }
      else cl.push({ v: pt.v, n: 1, rows: new Set([pt.li]) });
    });
    const need = Math.max(2, Math.round(rowsCount * 0.15));
    return cl.filter(c => c.rows.size >= need).map(c => c.v);
  };
  const byLeft = clusterBy(it => it.x);
  const byCenter = clusterBy(it => it.x + Math.max(it.w, 1) / 2);
  let anchors = (byLeft.length >= byCenter.length ? byLeft : byCenter);
  let pickPos = (byLeft.length >= byCenter.length) ? (it => it.x) : (it => it.x + Math.max(it.w, 1) / 2);

  let colOf;
  const usingAnchors = anchors.length >= 3;
  if (usingAnchors) {
    colOf = it => {
      const v = pickPos(it);
      let best = 0, bd = Infinity;
      anchors.forEach((a, i) => { const d = Math.abs(v - a); if (d < bd) { bd = d; best = i; } });
      return best;
    };
  } else {
    // そろった位置が見つからないときは、文字の占める区間の切れ目で列を切る
    const spans = src.map(it => [it.x, it.x + Math.max(it.w, 1)]).sort((a, b) => a[0] - b[0]);
    const blocks = [];
    spans.forEach(sp => {
      const last = blocks[blocks.length - 1];
      if (last && sp[0] <= last[1] + H * 0.4) last[1] = Math.max(last[1], sp[1]);
      else blocks.push([sp[0], sp[1]]);
    });
    anchors = blocks;
    colOf = it => {
      const c = it.x + Math.max(it.w, 1) / 2;
      let best = 0, bd = Infinity;
      blocks.forEach((b, i) => {
        const d = (c >= b[0] && c <= b[1]) ? 0 : Math.min(Math.abs(c - b[0]), Math.abs(c - b[1]));
        if (d < bd) { bd = d; best = i; }
      });
      return best;
    };
  }
  const colCount = anchors.length;

  const rows = [], ys = [];
  lines.forEach(line => {
    // まず各文字を列に割り当てる（左から順に見て、すき間なく続く文字は同じセルの続きとみなす）
    const bucket = [];
    let prevRight = -1e9, prevCol = 0, prevY = null;
    line.items.slice().sort((a, b) => a.x - b.x).forEach(it => {
      const own = colOf(it);
      const onAnchor = usingAnchors && Math.abs(pickPos(it) - anchors[own]) <= H * 0.35;
      const sameLevel = prevY !== null && Math.abs(it.y - prevY) <= H * 0.3;
      const cont = !onAnchor && sameLevel && it.x <= prevRight + H * 0.15;
      const i = cont ? prevCol : own;
      (bucket[i] = bucket[i] || []).push(it);
      prevRight = it.x + Math.max(it.w, 1); prevCol = i; prevY = it.y;
    });
    // セルの中身を組み立てる。セル内で上下2段になっている場合は上の段から読む
    const cells = [];
    for (let i = 0; i < colCount; i++) {
      const list = bucket[i];
      if (!list) { cells[i] = ''; continue; }
      list.sort((a, b) => (b.y - a.y) || (a.x - b.x));
      let text = '', lastY = null;
      list.forEach(it => {
        if (lastY === null) text = it.s;
        else if (Math.abs(it.y - lastY) <= H * 0.3) text += it.s;   // 同じ段の続き
        else text = joinCellText(text, it.s);                       // 折り返した次の段
        lastY = it.y;
      });
      cells[i] = text.trim();
    }
    rows.push(cells);
    ys.push(line.y);
  });
  return { grid: rows, ys: ys };
}

/* ---------- グリッド → シフト ---------- */
function scoreDateColumn(grid, c) {
  const entries = [];
  let mdCount = 0;
  for (let r = 0; r < grid.length; r++) {
    const d = readDate(grid[r][c]);
    if (!d) continue;
    if (d.mode !== 'day') mdCount++;
    entries.push({ r: r, d: d });
  }
  if (entries.length < 3) return { score: 0 };
  // 「日のみ」の列は、1〜31が増えていく並びのときだけ日付とみなす
  if (!mdCount) {
    let ok = true;
    for (let i = 1; i < entries.length; i++) if (entries[i].d.d <= entries[i - 1].d.d) { ok = false; break; }
    if (!ok || entries.length < 10) return { score: 0 };
  }
  return { score: entries.length + mdCount * 0.5, entries: entries, col: c };
}

function analyzeGrid(grid, fallbackYM, ys, keepNames) {
  const keep = keepNames || [];
  const width = grid.reduce((a, r) => Math.max(a, r.length), 0);
  let best = { score: 0 };
  for (let c = 0; c < width; c++) {
    const s = scoreDateColumn(grid, c);
    if (s.score > best.score) best = s;
  }
  if (!best.score) return null;

  const dateCol = best.col, firstRow = best.entries[0].r;

  // 見出し行：「日付」がある行、なければ日付行の直前で中身のある行
  let headerRow = -1;
  for (let r = 0; r < grid.length; r++) {
    if (String(grid[r][dateCol] || '').trim() === '日付') { headerRow = r; break; }
  }
  if (headerRow < 0) {
    for (let r = firstRow - 1; r >= 0; r--) {
      if (grid[r].some((t, i) => i !== dateCol && String(t).trim() !== '')) { headerRow = r; break; }
    }
  }
  if (headerRow < 0) headerRow = Math.max(0, firstRow - 1);

  const dateRows = best.entries.map(e => e.r);
  const isWeekdayCol = c => {
    let n = 0, w = 0;
    dateRows.forEach(r => { const t = String(grid[r][c] || '').trim(); if (t) { n++; if (WD_RE.test(t)) w++; } });
    return n >= 3 && w / n >= 0.6;
  };

  const staff = [];
  let openCol = -1;
  for (let c = 0; c < width; c++) {
    if (c === dateCol) continue;
    const h = String(grid[headerRow][c] || '').trim();
    if (h === '開館時間') { openCol = c; continue; }
    if (!h || SKIP_HEADERS.indexOf(h) >= 0 || SKIP_HEADER_RE.test(h) || readDate(h) || WD_RE.test(h)) continue;
    if (isWeekdayCol(c)) continue;
    if (staff.some(s => s.name === h)) continue;
    // 時間帯が1つも入っていない列はスタッフ列とみなさない（タイトルの誤検出よけ）。
    // ただし職員として指定された名前の列は、数字だけの記入でも残す。
    if (keep.indexOf(h) < 0) {
      const hasTime = dateRows.some(r => parseRanges(grid[r][c]).length > 0 || String(grid[r][c] || '').indexOf('終日') >= 0);
      if (!hasTime) continue;
    }
    staff.push({ name: h, col: c });
  }
  if (!staff.length) return null;

  const allText = grid.map(r => r.join(' ')).join(' ');
  const ym = detectYearMonth(allText) || {};
  const year = ym.y || (fallbackYM && fallbackYM.y) || new Date().getFullYear();
  const month = ym.m || (fallbackYM && fallbackYM.m) || (new Date().getMonth() + 1);
  let guessed = false;

  const days = best.entries.map((e, i) => {
    const dd = e.d;
    let y = dd.y, m = dd.m, d = dd.d;
    if (dd.mode !== 'full') { guessed = true; if (m === undefined) m = month; y = year; }
    let open = null, close = null;
    if (openCol >= 0) {
      const oc = parseRanges(grid[e.r][openCol]);
      if (oc.length) { open = oc[0][0]; close = oc[0][1]; }
    }
    const cells = {};
    const nextRow = (best.entries[i + 1] ? best.entries[i + 1].r : grid.length);
    const stop = Math.min(nextRow, e.r + 4);
    staff.forEach(s => {
      let t = String(grid[e.r][s.col] || '').trim();
      for (let rr = e.r + 1; rr < stop; rr++) {
        const x = String(grid[rr][s.col] || '').trim();
        if (x) t = joinCellText(t, x);
      }
      if (t) cells[s.name] = t;
    });
    return { y: y, m: m, d: d, wd: new Date(Date.UTC(y, m - 1, d)).getUTCDay(), open: open, close: close, cells: cells, y0: ys ? ys[e.r] : null };
  });

  const filled = days.filter(day => Object.keys(day.cells).some(n => parseCell(day.cells[n], day.open, day.close, 9).length));
  return {
    score: best.score + filled.length,
    staffNames: staff.map(s => s.name),
    days: days,
    guessedMonth: guessed,
    filledDays: filled.length
  };
}

/** グリッド（縦横どちらでも可）を解析する */
function parseGrid(grid, fallbackYM, ys, keepNames) {
  const a = analyzeGrid(grid, fallbackYM, ys, keepNames);
  const b = analyzeGrid(transpose(grid), fallbackYM, null, keepNames);
  if (!a && !b) return null;
  if (!a) return b;
  if (!b) return a;
  return (b.score > a.score) ? b : a;
}

/** 複数ページ・複数シートの結果をまとめる */
function mergeParsed(list) {
  const byDate = {}, order = [], names = [];
  let guessed = false;
  list.forEach(p => {
    if (!p) return;
    if (p.guessedMonth) guessed = true;
    p.staffNames.forEach(n => { if (names.indexOf(n) < 0) names.push(n); });
    p.days.forEach(day => {
      const k = day.y + '-' + day.m + '-' + day.d;
      if (!byDate[k]) { byDate[k] = { y: day.y, m: day.m, d: day.d, wd: day.wd, open: day.open, close: day.close, cells: {} }; order.push(k); }
      const t = byDate[k];
      if (day.open != null) { t.open = day.open; t.close = day.close; }
      Object.keys(day.cells).forEach(n => { if (!t.cells[n]) t.cells[n] = day.cells[n]; });
    });
  });
  if (!order.length) return null;
  const days = order.map(k => byDate[k]).sort((x, y) => (x.y - y.y) || (x.m - y.m) || (x.d - y.d));
  return { staffNames: names, days: days, guessedMonth: guessed };
}

/** セル内で折り返された行を、上の行に取り込む
 *  （左端の見出し列が空で、文字の入っている列が上の行の一部に収まっている行が折り返し） */
function mergeWrappedRows(page) {
  const grid = [], ys = [];
  page.grid.forEach((row, i) => {
    const prev = grid[grid.length - 1];
    const cols = row.map((t, c) => (t ? c : -1)).filter(c => c >= 0);
    const isWrap = prev && cols.length > 0 && !row[0] &&
      !row.some(t => readDate(t)) &&
      cols.every(c => prev[c] && prev[c] !== '');
    if (isWrap) {
      cols.forEach(c => { prev[c] = joinCellText(prev[c], row[c]); });
    } else { grid.push(row.slice()); ys.push(page.ys[i]); }
  });
  return { grid: grid, ys: ys };
}

/** 複数ページのPDFを、行のy座標をたよりに1枚の表につなぎ直す
 *  （表が横に分割されて印刷されている場合は横につなぎ、そうでなければ下に足す） */
function stitchPages(pages) {
  if (!pages.length) return { grid: [], ys: [] };
  const base = { grid: pages[0].grid.map(r => r.slice()), ys: pages[0].ys.slice() };
  for (let i = 1; i < pages.length; i++) {
    const pg = pages[i];
    if (!pg.ys.length) continue;
    let match = 0;
    pg.ys.forEach(y => { if (base.ys.some(by => Math.abs(by - y) <= 3)) match++; });
    const width = base.grid.reduce((a, r) => Math.max(a, r.length), 0);
    if (match / pg.ys.length >= 0.6) {
      pg.grid.forEach((row, ri) => {
        let bi = -1, bd = 3.5;
        base.ys.forEach((by, k) => { const d = Math.abs(by - pg.ys[ri]); if (d < bd) { bd = d; bi = k; } });
        if (bi < 0) {
          const nr = [];
          for (let c = 0; c < width; c++) nr.push('');
          row.forEach((t, c) => { nr[width + c] = t; });
          base.grid.push(nr); base.ys.push(pg.ys[ri]);
        } else {
          const r2 = base.grid[bi];
          while (r2.length < width) r2.push('');
          row.forEach((t, c) => { r2[width + c] = t; });
        }
      });
    } else {
      pg.grid.forEach((row, ri) => { base.grid.push(row.slice()); base.ys.push(pg.ys[ri]); });
    }
  }
  const W = base.grid.reduce((a, r) => Math.max(a, r.length), 0);
  base.grid.forEach(r => { for (let c = 0; c < W; c++) if (r[c] === undefined) r[c] = ''; });
  return base;
}

/* ---------- アルバイトの列割り当て ---------- */
/** アルバイトを決まった数の列に振り分ける。
 *  勤務が重ならない人どうしは同じ列にまとめ、列は「人数の少ない列を左から」選ぶ
 *  （校舎で実際に使われている並べ方に合わせている）。 */
function packStaff(names, shifts, colCount) {
  const cols = [];
  for (let i = 0; i < colCount; i++) cols.push({ names: [], end: -1 });
  const sorted = names.slice().sort((a, b) =>
    (shifts[a][0][0] - shifts[b][0][0]) || a.localeCompare(b, 'ja'));
  const overflow = [];
  sorted.forEach(n => {
    const rg = shifts[n];
    const start = rg[0][0], end = rg[rg.length - 1][1];
    const free = cols.filter(c => c.end <= start);
    let target;
    if (free.length) {
      const least = Math.min.apply(null, free.map(c => c.names.length));
      target = free.filter(c => c.names.length === least)[0];
    } else {
      // どの列とも重なるとき（同時に4人以上）は、勤務終了が同じ（いちばん近い）列に相乗りさせる。
      // 終了が同じなら列の下がそろうので、表として読みやすい。
      target = cols.slice().sort((a, b) =>
        (Math.abs(a.end - end) - Math.abs(b.end - end)) ||
        (a.names.length - b.names.length) ||
        (cols.indexOf(a) - cols.indexOf(b)))[0];
      overflow.push(n);
    }
    target.names.push(n);
    target.end = Math.max(target.end, end);
  });
  return { cols: cols.map(c => c.names), overflow: overflow };
}

/** 勤務開始時刻が入る行（表は9:00〜21:30の30分刻み） */
function rowOfTime(min) {
  const i = Math.floor((min - DAY_START) / SLOT);
  return FIRST_ROW + Math.min(Math.max(i, 0), SLOTS - 1);
}

/** 複数人ぶんの勤務時間帯をひとつにまとめる */
function unionRanges(list) {
  const all = [];
  list.forEach(rg => rg.forEach(r => all.push(r.slice())));
  all.sort((a, b) => a[0] - b[0]);
  const out = [];
  all.forEach(r => {
    const last = out[out.length - 1];
    if (last && r[0] <= last[1]) last[1] = Math.max(last[1], r[1]);
    else out.push(r.slice());
  });
  return out;
}


/* ---------- 「今週のAA」出勤予定表 ---------- */
const POSTER_ROWS = 7;      // 1ページの日数
const POSTER_LINES = 3;     // 1日あたりの行数

const hhmm = m => Math.floor(m / 60) + ':' + String(m % 60).padStart(2, '0');

/** その日のアルバイトを「同じ時間帯どうし」でまとめる */
function posterGroups(day, fixed, order) {
  const items = [];
  Object.keys(day.staff).forEach(n => {
    if (fixed.indexOf(n) >= 0) return;
    day.staff[n].forEach(r => items.push({ n: n, a: r[0], b: r[1] }));
  });
  const map = {};
  items.forEach(it => {
    const k = it.a + '-' + it.b;
    if (!map[k]) map[k] = { a: it.a, b: it.b, names: [] };
    map[k].names.push(it.n);
  });
  const rank = n => { const i = (order || []).indexOf(n); return i < 0 ? 999 : i; };
  const rs = day.rawStart || {};
  const start = n => (rs[n] === undefined ? 0 : rs[n]);
  return Object.keys(map).map(k => map[k])
    .sort((x, y) => (x.a - y.a) || (x.b - y.b))
    .map(g => {
      // 同じ枠の中は、実際の出勤が早い人から並べる（表の列順は同着のときの目安）
      g.names.sort((a, b) => (start(a) - start(b)) || (rank(a) - rank(b)) || a.localeCompare(b, 'ja'));
      return g;
    });
}

/** 時間帯のまとまりを3行に配り分ける（多い分は真ん中の行から足す） */
function posterLines(groups) {
  const n = groups.length;
  if (n <= POSTER_LINES) return groups.map(g => [g]);
  const counts = [1, 1, 1], seq = [1, 2, 0];
  for (let i = 0, rest = n - POSTER_LINES; rest > 0; i++, rest--) counts[seq[i % 3]]++;
  const lines = []; let p = 0;
  counts.forEach(c => { lines.push(groups.slice(p, p + c)); p += c; });
  return lines;
}

function escHtml(s) {
  return String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}

/** 「今週のAA」の印刷用HTMLを組み立てる */
function buildPoster(days, fixed, order, month) {
  const pages = [];
  for (let i = 0; i < days.length; i += POSTER_ROWS) pages.push(days.slice(i, i + POSTER_ROWS));

  const rowHtml = (day, idx) => {
    const alt = idx % 2 ? ' alt' : '';
    if (!day) return `<tr class="row${alt}"><td class="day"></td><td class="lines"></td></tr>`;
    const cls = day.wd === 0 ? ' sun' : (day.wd === 6 ? ' sat' : '');
    const lines = posterLines(posterGroups(day, fixed, order)).map(gs =>
      '<div class="line">' + gs.map(g =>
        `<span class="grp"><span class="t">${hhmm(g.a)}~${hhmm(g.b)}</span><span class="nm">${g.names.map(escHtml).join('　')}</span></span>`
      ).join('') + '</div>').join('');
    return `<tr class="row${alt}">
      <td class="day${cls}"><span class="num">${day.d}</span><span class="wd">${WD[day.wd]}</span></td>
      <td class="lines"><div class="box" contenteditable="true">${lines}</div></td></tr>`;
  };

  const body = pages.map(page => {
    const rows = [];
    for (let i = 0; i < POSTER_ROWS; i++) rows.push(rowHtml(page[i], i));
    return `<div class="page">
      <div class="head"><span class="mon">${month}</span><span class="tsuki">月</span><span class="title">今 週 の Ａ Ａ</span></div>
      <table class="card">${rows.join('')}</table>
      <div class="foot">都合によりAAが変更になる場合があります。</div>
    </div>`;
  }).join('');

  return `<!DOCTYPE html><html lang="ja"><head><meta charset="utf-8">
<title>${month}月 今週のAA</title>
<style>
  @page { size:A4 portrait; margin:0; }
  html,body { margin:0; padding:0; background:#7E7E7E;
    font-family:"Hiragino Kaku Gothic ProN","Yu Gothic","Meiryo","Noto Sans JP",sans-serif;
    -webkit-print-color-adjust:exact; print-color-adjust:exact; }
  .page { width:210mm; height:296.6mm; background:#F98D87; padding:7mm 8mm 0; box-sizing:border-box;
    page-break-after:always; overflow:hidden; margin:0 auto; }
  .page:last-child { page-break-after:auto; }
  .head { height:30mm; text-align:center; color:#111; white-space:nowrap; }
  .mon { font-size:21mm; font-weight:900; line-height:30mm; vertical-align:middle; }
  .tsuki { font-size:5.5mm; font-weight:700; margin:0 7mm 0 2mm; vertical-align:middle; }
  .title { font-size:12mm; font-weight:900; letter-spacing:3mm; vertical-align:middle; }
  .card { width:100%; height:236mm; background:#fff; border-collapse:collapse; table-layout:fixed;
    border-top:1.4mm solid #111; border-bottom:1.4mm solid #111; }
  .row { height:33.7mm; }
  .row td { border-bottom:.35mm solid #111; }
  .row:last-child td { border-bottom:none; }
  .row.alt td { background:#FCF6D4; }
  .day { width:31mm; text-align:center; vertical-align:middle; border-right:.35mm solid #111; color:#111; }
  .num { font-size:11mm; font-weight:900; vertical-align:middle; margin-right:2.5mm; }
  .wd { display:inline-block; width:9mm; height:9mm; line-height:8.4mm; border:.35mm solid currentColor;
    border-radius:50%; font-size:4.8mm; font-weight:700; vertical-align:middle; }
  .sat { color:#3B4FD8; } .sun { color:#E8332A; }
  .lines { vertical-align:middle; padding:2mm 4mm; }
  .box { outline:none; }
  .line { white-space:nowrap; margin-bottom:1.8mm; }
  .line:last-child { margin-bottom:0; }
  .grp { display:inline-block; margin-right:9mm; }
  .grp:last-child { margin-right:0; }
  .t { font-size:5.2mm; font-weight:700; }
  .nm { font-size:5.2mm; font-weight:700; margin-left:5mm; }
  .foot { text-align:center; font-size:3.6mm; font-weight:700; color:#111; line-height:12mm; }
  @media screen { body { padding:6mm 0; } .page { box-shadow:0 2px 14px rgba(0,0,0,.4); margin-bottom:6mm; } }
</style></head><body>${body}</body></html>`;
}


/* ---------- Excel 書き出し ---------- */
const cloneStyle = cell => JSON.parse(JSON.stringify(cell.style || {}));
const solid = argb => ({ type: 'pattern', pattern: 'solid', fgColor: { argb: argb }, bgColor: { argb: argb } });

/** 1日分のシートを、原版の書式を複製して作る */
function buildDaySheet(wb, tpl, day, fixed) {
  const ws = wb.addWorksheet(String(day.d));
  const part = Object.keys(day.staff).filter(n => fixed.indexOf(n) < 0);
  const packed = packStaff(part, day.staff, MIN_PT).cols;
  // 列の中身：職員は1人ずつ、アルバイトは同居あり
  const slots = fixed.map(n => (day.staff[n] ? [n] : [n]))
    .concat(packed.map(list => list.slice()));
  const nSlots = slots.length;

  const map = [{ dst: 1, src: SRC.time }];
  for (let i = 0; i < nSlots; i++) {
    map.push({ dst: 2 + 2 * i, src: SRC.narrow, slot: i });
    map.push({ dst: 3 + 2 * i, src: SRC.wide, slot: i });
  }
  const spacer = 2 + 2 * nSlots;
  map.push({ dst: spacer, src: SRC.spacer });
  map.push({ dst: spacer + 1, src: SRC.b1 });
  map.push({ dst: spacer + 2, src: SRC.b2 });
  map.push({ dst: spacer + 3, src: SRC.tail });

  map.forEach(mp => {
    const srcCol = tpl.getColumn(mp.src);
    if (srcCol && srcCol.width) ws.getColumn(mp.dst).width = srcCol.width;
    for (let r = 1; r <= TPL_LAST_ROW; r++) {
      const s = tpl.getCell(r, mp.src), t = ws.getCell(r, mp.dst);
      t.style = cloneStyle(s);
      if (mp.slot === undefined && s.value !== null && s.value !== undefined) t.value = s.value;
    }
  });
  for (let r = 1; r <= TPL_LAST_ROW; r++) {
    const h = tpl.getRow(r).height;
    if (h) ws.getRow(r).height = h;
  }

  ws.getCell(1, spacer + 1).value = new Date(Date.UTC(day.y, day.m - 1, day.d));

  for (let i = 0; i < nSlots; i++) {
    const nCol = 2 + 2 * i, wCol = nCol + 1;
    const list = slots[i] || [];
    const ranges = unionRanges(list.map(n => day.staff[n] || []));
    ws.getCell(NAME_ROW, nCol).value = list.length ? list[0] : null;
    for (let k = 0; k < SLOTS; k++) {
      const fill = solid(isWorking(ranges, k) ? WHITE : GRAY);
      ws.getCell(FIRST_ROW + k, nCol).fill = fill;
      ws.getCell(FIRST_ROW + k, wCol).fill = fill;
    }
    // 2人目以降は、勤務開始の行に名前を書き入れる（校舎での書き方に合わせる）
    for (let k = 1; k < list.length; k++) {
      const rg = day.staff[list[k]];
      if (!rg || !rg.length) continue;
      const cell = ws.getCell(rowOfTime(rg[0][0]), nCol);
      cell.value = cell.value ? cell.value + '・' + list[k] : list[k];
    }
  }
  for (let i = 0; i < nSlots; i++) ws.mergeCells(NAME_ROW, 2 + 2 * i, NAME_ROW, 3 + 2 * i);
  ws.mergeCells(1, spacer + 1, 1, spacer + 2);

  if (tpl.views) ws.views = JSON.parse(JSON.stringify(tpl.views));
  if (tpl.pageSetup) ws.pageSetup = JSON.parse(JSON.stringify(tpl.pageSetup));
  return ws;
}

async function buildWorkbook(ExcelJSRef, templateBuf, days, fixed) {
  const wb = new ExcelJSRef.Workbook();
  await wb.xlsx.load(templateBuf);
  let tpl = wb.getWorksheet('原版(触らない)');
  if (!tpl) tpl = wb.worksheets.filter(w => w.name.indexOf('原版') === 0)[0];
  if (!tpl) throw new Error('テンプレートに「原版(触らない)」シートが見つかりません。');
  wb.worksheets.filter(w => /^\d+$/.test(w.name)).forEach(w => wb.removeWorksheet(w.id));
  days.forEach(day => buildDaySheet(wb, tpl, day, fixed));
  return await wb.xlsx.writeBuffer();
}
// ===CORE END===

/* ======================= 画面 ======================= */
const KEY_TPL = 'meeting-schedule:template';
const KEY_CFG = 'meeting-schedule:config';

const state = {
  tplBuf: null, tplName: '', tplSaved: false,
  shiftName: '', raw: null,            // 読み込んだファイルのグリッド（再解析用）
  parsed: null,                        // {staffNames, days:[{y,m,d,wd,open,close,cells}]}
  months: [], month: '',
  fixed: ['平野', '小西'],
  staffInput: {},                      // 'Y-M-D' -> {職員名: 'テキスト'}
  pattern: {},                         // 曜日 -> {職員名: 'テキスト'}
  ym: null,                            // 年月を手で指定した場合
  hold: 9,                             // 職員の拘束時間（時間）
};

const $ = id => document.getElementById(id);
const pad = n => String(n).padStart(2, '0');
const keyOf = d => `${d.y}-${pad(d.m)}-${pad(d.d)}`;
const esc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

function b64encode(buf) {
  const bytes = new Uint8Array(buf); let s = '';
  for (let i = 0; i < bytes.length; i += 8192) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 8192));
  return btoa(s);
}
function b64decode(str) {
  const bin = atob(str), bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return bytes.buffer;
}
function readFile(file) {
  return new Promise((res, rej) => {
    const fr = new FileReader();
    fr.onload = () => res(fr.result);
    fr.onerror = () => rej(new Error('ファイルを読み込めませんでした。'));
    fr.readAsArrayBuffer(file);
  });
}
function setStatus(msg, isErr) {
  $('status').textContent = msg;
  $('status').className = 'status' + (isErr ? ' err' : '');
}

function dropZone(el, label, sub, accept, onFile) {
  el.innerHTML = `<div class="drop" tabindex="0" role="button"><strong>${label}</strong><em>${sub}</em></div>`;
  const zone = el.firstElementChild;
  const input = document.createElement('input');
  input.type = 'file'; input.accept = accept; input.style.display = 'none';
  el.appendChild(input);
  zone.addEventListener('click', () => input.click());
  zone.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); input.click(); } });
  input.addEventListener('change', () => { if (input.files[0]) onFile(input.files[0]); input.value = ''; });
  ['dragenter', 'dragover'].forEach(t => zone.addEventListener(t, e => { e.preventDefault(); zone.classList.add('hot'); }));
  ['dragleave', 'drop'].forEach(t => zone.addEventListener(t, e => { e.preventDefault(); zone.classList.remove('hot'); }));
  zone.addEventListener('drop', e => { const f = e.dataTransfer.files[0]; if (f) onFile(f); });
}

/* ---------- テンプレート ---------- */
function renderTemplate() {
  const zone = $('tplZone');
  if (state.tplBuf) {
    zone.innerHTML = `<div class="filed"><div class="nm"><b>${esc(state.tplName || '保存済みのテンプレート')}</b>
      <small>${state.tplSaved ? 'この端末に保存済み・次回から読み込み不要' : '読み込み済み'}</small></div>
      <button class="btn-ghost" id="tplChange">変更</button></div>`;
    $('tplChange').addEventListener('click', () => { state.tplBuf = null; renderTemplate(); refresh(); });
  } else {
    dropZone(zone, 'テンプレートのExcelを選ぶ', 'ドラッグ＆ドロップでも読み込めます', '.xlsx,.xlsm', async file => {
      try {
        const buf = await readFile(file);
        const wb = new ExcelJS.Workbook(); await wb.xlsx.load(buf);
        if (!wb.getWorksheet('原版(触らない)') && !wb.worksheets.some(w => w.name.indexOf('原版') === 0)) {
          throw new Error('「原版(触らない)」シートが見当たりません。面談スケジュール表のファイルを選んでください。');
        }
        state.tplBuf = buf; state.tplName = file.name; state.tplSaved = false;
        try {
          await window.storage.set(KEY_TPL, JSON.stringify({ name: file.name, data: b64encode(buf) }));
          state.tplSaved = true;
        } catch (e) { }
        renderTemplate(); refresh(); setStatus('テンプレートを読み込みました。');
      } catch (e) { setStatus(e.message, true); }
    });
  }
}

/* ---------- シフト表の読み込み ---------- */
async function gridsFromExcel(buf) {
  const wb = new ExcelJS.Workbook(); await wb.xlsx.load(buf);
  return { kind: 'excel', sheets: wb.worksheets.map(ws => ({ name: ws.name, grid: gridFromWorksheet(ws) })) };
}

async function gridsFromPdf(buf) {
  const lib = window.pdfjsLib || window['pdfjs-dist/build/pdf'];
  if (!lib) throw new Error('PDFの読み取り機能を読み込めませんでした。通信環境を確認して開き直してください。');
  try { lib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js'; } catch (e) { }
  const doc = await lib.getDocument({ data: new Uint8Array(buf) }).promise;
  const pages = [];
  for (let p = 1; p <= doc.numPages; p++) {
    const page = await doc.getPage(p);
    const tc = await page.getTextContent();
    const items = tc.items.filter(i => i.str && i.str.trim()).map(i => ({
      s: i.str, x: i.transform[4], y: i.transform[5], w: i.width || 0, h: Math.abs(i.transform[3]) || 10
    }));
    if (items.length) pages.push(gridFromItems(items));
  }
  if (!pages.length) throw new Error('PDFから文字を読み取れませんでした。紙をスキャンしただけのPDFには対応していません。');
  const st = mergeWrappedRows(stitchPages(pages));
  return { kind: 'pdf', sheets: [{ name: `${doc.numPages}ページ`, grid: st.grid, ys: st.ys }] };
}

function reparse() {
  if (!state.raw) return null;
  const list = state.raw.sheets.map(sh => {
    try { return parseGrid(sh.grid, state.ym, sh.ys, state.fixed); } catch (e) { return null; }
  });
  const named = state.raw.sheets.map((sh, i) => ({ sh: sh, p: list[i] })).filter(x => x.p);
  const merged = mergeParsed(named.map(x => x.p));
  if (merged) merged.sheetName = named.map(x => x.sh.name).join('・');
  return merged;
}

async function loadShift(file) {
  const buf = await readFile(file);
  const isPdf = /\.pdf$/i.test(file.name);
  state.raw = isPdf ? await gridsFromPdf(buf) : await gridsFromExcel(buf);
  state.ym = null;
  const parsed = reparse();
  if (!parsed) throw new Error('日付とスタッフ名の表を見つけられませんでした。別の形式のファイルなら送ってもらえれば対応します。');
  state.parsed = parsed;
  state.shiftName = file.name + '（' + parsed.sheetName + '）';
  state.staffInput = {};

  const count = {};
  parsed.days.forEach(d => { const k = `${d.y}-${pad(d.m)}`; count[k] = (count[k] || 0) + 1; });
  state.months = Object.keys(count).sort();
  state.month = state.months.slice().sort((a, b) => count[b] - count[a])[0];
  $('monthField').style.display = state.months.length > 1 ? 'block' : 'none';
  $('monthSel').innerHTML = state.months.map(m =>
    `<option value="${m}" ${m === state.month ? 'selected' : ''}>${m.split('-')[0]}年${Number(m.split('-')[1])}月（${count[m]}日）</option>`).join('');

  $('ymFix').style.display = parsed.guessedMonth ? 'block' : 'none';
  if (parsed.guessedMonth) {
    const [y, m] = state.month.split('-');
    $('ymY').value = y; $('ymM').value = Number(m);
  }
  seedStaffFromShift();
}

function renderShift() {
  const zone = $('shiftZone');
  if (state.shiftName) {
    const p = state.parsed;
    zone.innerHTML = `<div class="filed"><div class="nm"><b>${esc(state.shiftName)}</b>
      <small>スタッフ${p.staffNames.length}名・${p.days.length}日分を読み取り</small></div>
      <button class="btn-ghost" id="shiftChange">変更</button></div>`;
    $('shiftChange').addEventListener('click', () => {
      state.shiftName = ''; state.raw = null; state.parsed = null; state.months = []; state.month = '';
      $('ymFix').style.display = 'none'; $('checkBox').style.display = 'none';
      renderShift(); renderStaff(); renderCheck(); renderPreview(); refresh();
    });
  } else {
    dropZone(zone, 'シフト表を選ぶ（Excel / PDF）', 'ドラッグ＆ドロップでも読み込めます', '.xlsx,.xlsm,.pdf', async file => {
      setStatus('読み取っています…');
      try {
        await loadShift(file);
        renderShift(); renderStaff(); renderCheck(); renderPreview(); refresh();
        setStatus('シフト表を読み取りました。内容を確認してから書き出してください。');
      } catch (e) { setStatus(e.message, true); }
    });
  }
}

/* ---------- 読み取り結果の確認 ---------- */
function monthDays() {
  if (!state.parsed) return [];
  return state.parsed.days.filter(d => `${d.y}-${pad(d.m)}` === state.month).sort((a, b) => a.d - b.d);
}
function partNames() {
  if (!state.parsed) return [];
  return state.parsed.staffNames.filter(n => state.fixed.indexOf(n) < 0);
}

function renderCheck() {
  const days = monthDays(), names = partNames();
  if (!days.length || !names.length) { $('checkBox').style.display = 'none'; return; }
  $('checkBox').style.display = 'block';
  $('cellHead').innerHTML = '<tr><th>日</th>' + names.map(n =>
    `<th>${esc(n)}<button data-del="${esc(n)}" title="この列を削除">×</button></th>`).join('') + '</tr>';
  $('cellBody').innerHTML = days.map(d => {
    const k = keyOf(d);
    return `<tr><td class="d${(d.wd === 0 || d.wd === 6) ? ' we' : ''}">${d.m}/${d.d}(${WD[d.wd]})</td>` +
      names.map(n => `<td><input type="text" data-k="${k}" data-nm="${esc(n)}" value="${esc(d.cells[n] || '')}"></td>`).join('') + '</tr>';
  }).join('');
  $('cellBody').querySelectorAll('input').forEach(inp => {
    inp.addEventListener('input', () => {
      const d = monthDays().find(x => keyOf(x) === inp.dataset.k);
      if (!d) return;
      const v = inp.value.trim();
      if (v) d.cells[inp.dataset.nm] = v; else delete d.cells[inp.dataset.nm];
      renderPreview();
    });
  });
  $('cellHead').querySelectorAll('button').forEach(b => {
    b.addEventListener('click', () => {
      const n = b.dataset.del;
      state.parsed.staffNames = state.parsed.staffNames.filter(x => x !== n);
      state.parsed.days.forEach(d => delete d.cells[n]);
      renderCheck(); renderPreview();
    });
  });
}

/* ---------- 職員シフト ---------- */
function seedStaffFromShift() {
  monthDays().forEach(d => {
    const k = keyOf(d);
    state.staffInput[k] = state.staffInput[k] || {};
    state.fixed.forEach(n => {
      if (state.staffInput[k][n] === undefined && d.cells[n]) state.staffInput[k][n] = d.cells[n];
    });
  });
}

function renderStaff() {
  const days = monthDays();
  if (!days.length) { $('staffCard').style.display = 'none'; return; }
  $('staffCard').style.display = 'block';
  const fixed = state.fixed;

  $('pattern').innerHTML = [1, 2, 3, 4, 5, 6, 0].map(w =>
    `<label><span>${WD[w]}曜</span>${fixed.map(n =>
      `<input type="text" data-wd="${w}" data-nm="${esc(n)}" placeholder="${esc(n)}" value="${esc((state.pattern[w] || {})[n] || '')}">`
    ).join('')}</label>`).join('');
  $('pattern').querySelectorAll('input').forEach(inp => {
    inp.addEventListener('input', () => {
      const w = inp.dataset.wd;
      state.pattern[w] = state.pattern[w] || {};
      state.pattern[w][inp.dataset.nm] = inp.value;
      saveConfig();
    });
  });

  $('staffHead').innerHTML = `<tr><th>日</th>${fixed.map(n => `<th>${esc(n)}</th>`).join('')}</tr>`;
  $('staffBody').innerHTML = days.map(d => {
    const k = keyOf(d), cur = state.staffInput[k] || {};
    return `<tr><td class="d${(d.wd === 0 || d.wd === 6) ? ' we' : ''}">${d.m}/${d.d}(${WD[d.wd]})</td>` +
      fixed.map(n => `<td><input type="text" data-k="${k}" data-nm="${esc(n)}" value="${esc(cur[n] || '')}" placeholder="例 12 / 14:00-21:00"></td>`).join('') + '</tr>';
  }).join('');
  $('staffBody').querySelectorAll('input').forEach(inp => {
    inp.addEventListener('input', () => {
      const k = inp.dataset.k;
      state.staffInput[k] = state.staffInput[k] || {};
      state.staffInput[k][inp.dataset.nm] = inp.value;
      renderPreview();
    });
  });
}

/* ---------- 生成対象 ---------- */
function resolvedDays() {
  return monthDays().map(d => {
    const staff = {}, rawStart = {};
    Object.keys(d.cells).forEach(n => {
      if (state.fixed.indexOf(n) >= 0) return;
      // アルバイトは開館時間で切らず、表の範囲（9:00〜22:00）に収めるだけにする
      const raw = parseCell(d.cells[n], d.open, d.close, null, false);
      if (!raw.length) return;
      const end = Math.min(d.close == null ? DAY_END : d.close, DAY_END);
      const rg = clipRanges(raw, DAY_START, end);
      if (rg.length) { staff[n] = rg; rawStart[n] = raw[0][0]; }
    });
    const inputs = state.staffInput[keyOf(d)] || {};
    state.fixed.forEach(n => {
      // 職員は数字だけの記入（勤務開始時刻）にも対応する
      const rg = parseCell(inputs[n], d.open, d.close, state.hold);
      if (rg.length) staff[n] = rg;
    });
    return { y: d.y, m: d.m, d: d.d, wd: d.wd, open: d.open, close: d.close, staff: staff, rawStart: rawStart };
  });
}

/* ---------- プレビュー ---------- */
function renderPreview() {
  const days = resolvedDays(), box = $('preview'), alerts = [];
  if (!days.length) {
    box.innerHTML = `<div class="empty"><b>まだ読み込んでいません</b>
      シフト表を読み込むと、1日ずつの出勤帯をここで確認できます。</div>`;
    $('legend').style.display = 'none'; $('previewCount').textContent = ''; $('alerts').innerHTML = '';
    $('previewTitle').textContent = 'プレビュー';
    return;
  }
  $('legend').style.display = 'flex';
  const parts = state.month.split('-');
  $('previewTitle').textContent = `${parts[0]}年${Number(parts[1])}月`;
  $('previewCount').textContent = `${days.length}日分 / ${days.length}シート`;

  let noPt = 0, over = 0;
  box.innerHTML = `<div class="scale"><span></span><span class="ticks">
      <i>9</i><i>10</i><i>11</i><i>12</i><i>13</i><i>14</i><i>15</i><i>16</i><i>17</i><i>18</i><i>19</i><i>20</i><i>21</i>
    </span></div>` + days.map(day => {
    const part = Object.keys(day.staff).filter(n => state.fixed.indexOf(n) < 0);
    if (!part.length) noPt++;
    const packed = packStaff(part, day.staff, MIN_PT);
    if (packed.overflow.length) over++;

    const line = (label, ranges, cls) => {
      const bands = Array.from({ length: SLOTS }, (_, k) => `<i class="${isWorking(ranges, k) ? 'on' : ''}"></i>`).join('');
      return `<div class="row"><span class="who ${cls}">${label}</span><span class="bands">${bands}</span></div>`;
    };
    const rows = state.fixed.map(n => {
      const rg = day.staff[n] || [];
      return line(esc(n) + (rg.length ? '' : '（休）'), rg, 'staff' + (rg.length ? '' : ' off'));
    }).concat(packed.cols.map(list => {
      if (!list.length) return line('<span style="color:#A6ADB6">空き</span>', [], 'off');
      const rg = unionRanges(list.map(n => day.staff[n] || []));
      return line(list.map(esc).join('・'), rg, '');
    })).join('');

    const flag = (!part.length || packed.overflow.length > 0);
    const note = !part.length ? '<div class="note">アルバイトの割当がありません</div>'
      : (packed.overflow.length ? `<div class="note">同じ時間に4名以上 → ${packed.overflow.map(esc).join('・')} を終了時刻の近い人と同じ列に入れました</div>` : '');
    return `<div class="day${flag ? ' alert' : ''}">
        <div class="daylabel"><b>${day.m}/${day.d}</b><small>${WD[day.wd]}曜</small></div>
        <div class="rows">${rows}${note}</div></div>`;
  }).join('');

  if (noPt) alerts.push(`アルバイトの割当がない日が ${noPt}日 あります。読み取り結果を確認してください。`);
  if (over) alerts.push(`同じ時間に4名以上いる日が ${over}日 あります。勤務終了が同じ（いちばん近い）人と同じ列にまとめています。`);
  $('alerts').innerHTML = alerts.map(a => `<li>${a}</li>`).join('');
}

/* ---------- 今週のAA ---------- */
function generatePoster() {
  const days = resolvedDays();
  if (!days.length) return;
  try {
    const order = (state.parsed && state.parsed.staffNames) || [];
    const html = buildPoster(days, state.fixed, order, Number(state.month.split('-')[1]));
    const blob = new Blob([html], { type: 'text/html;charset=utf-8' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `今週のAA_${state.month}.html`;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
    setStatus(`今週のAA_${state.month}.html を書き出しました。ブラウザで開いて印刷するとPDFになります。`);
  } catch (e) {
    setStatus('書き出せませんでした：' + e.message, true);
  }
}

/* ---------- 書き出し ---------- */
async function generate() {
  const days = resolvedDays();
  if (!state.tplBuf || !days.length) return;
  $('genBtn').disabled = true;
  setStatus('書き出しています…');
  try {
    const buf = await buildWorkbook(ExcelJS, state.tplBuf, days, state.fixed);
    const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `面談スケジュール表_${state.month}.xlsx`;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
    setStatus(`面談スケジュール表_${state.month}.xlsx を書き出しました（${days.length}シート）。`);
  } catch (e) {
    setStatus('書き出せませんでした：' + e.message, true);
  }
  $('genBtn').disabled = false;
}

function refresh() {
  const ok = !!(state.tplBuf && monthDays().length);
  $('genBtn').disabled = !ok;
  $('posterBtn').disabled = !monthDays().length;
  if (ok) setStatus('準備できました。書き出す前にプレビューで出勤帯を確認してください。');
}

async function saveConfig() {
  try { await window.storage.set(KEY_CFG, JSON.stringify({ fixed: state.fixed, pattern: state.pattern, hold: state.hold })); } catch (e) { }
}

function redrawAll() { renderStaff(); renderCheck(); renderPreview(); refresh(); }

/* ---------- 起動 ---------- */
(async function init() {
  renderTemplate(); renderShift(); renderPreview();

  $('monthSel').addEventListener('change', e => { state.month = e.target.value; seedStaffFromShift(); redrawAll(); });
  $('fixedNames').addEventListener('change', e => {
    state.fixed = e.target.value.split(/[,、]/).map(s => s.trim()).filter(Boolean);
    if (state.raw) {                       // 職員名が変わると読み取る列も変わるので読み直す
      const p = reparse();
      if (p) { state.parsed = p; state.staffInput = {}; }
    }
    seedStaffFromShift(); renderShift(); redrawAll(); saveConfig();
  });
  $('holdHours').addEventListener('change', e => {
    const v = parseFloat(e.target.value);
    if (v > 0) { state.hold = v; renderPreview(); saveConfig(); }
  });
  ['ymY', 'ymM'].forEach(id => $(id).addEventListener('change', () => {
    const y = +$('ymY').value, m = +$('ymM').value;
    if (!y || !m) return;
    state.ym = { y: y, m: m };
    const p = reparse();
    if (!p) { setStatus('指定した年月では読み取れませんでした。', true); return; }
    state.parsed = p; state.staffInput = {};
    state.month = `${y}-${pad(m)}`;
    state.months = [state.month];
    $('monthField').style.display = 'none';
    seedStaffFromShift(); renderShift(); redrawAll();
  }));
  $('applyPattern').addEventListener('click', () => {
    monthDays().forEach(d => {
      const k = keyOf(d);
      state.staffInput[k] = state.staffInput[k] || {};
      state.fixed.forEach(n => {
        const p = (state.pattern[d.wd] || {})[n];
        if (p && !state.staffInput[k][n]) state.staffInput[k][n] = p;
      });
    });
    renderStaff(); renderPreview();
  });
  $('clearStaff').addEventListener('click', () => {
    monthDays().forEach(d => { state.staffInput[keyOf(d)] = {}; });
    renderStaff(); renderPreview();
  });
  $('genBtn').addEventListener('click', generate);
  $('posterBtn').addEventListener('click', generatePoster);

  try {
    const cfg = await window.storage.get(KEY_CFG);
    if (cfg) {
      const c = JSON.parse(cfg.value);
      if (c.fixed && c.fixed.length) { state.fixed = c.fixed; $('fixedNames').value = c.fixed.join(','); }
      if (c.pattern) state.pattern = c.pattern;
      if (c.hold) { state.hold = c.hold; $('holdHours').value = c.hold; }
    }
  } catch (e) { }
  try {
    const saved = await window.storage.get(KEY_TPL);
    if (saved) {
      const obj = JSON.parse(saved.value);
      state.tplBuf = b64decode(obj.data); state.tplName = obj.name; state.tplSaved = true;
      renderTemplate(); refresh();
    }
  } catch (e) { }
})();

})();

(function(){

// ===CORE START===
const FIRST_ROW = 3, LAST_ROW = 28, DAY_START = 9 * 60, SLOT = 30, SLOTS = 26;
const NAME_ROW = 2, MIN_PT = 3;
const DAY_END = 9 * 60 + 26 * 30;   // 表のいちばん下（22:00）
const SRC = { time: 1, narrow: 2, wide: 3, spacer: 12, b1: 13, b2: 14, tail: 15 };
const TPL_LAST_ROW = 33;
const GRAY = 'FFD9D9D9', WHITE = 'FFFFFFFF';
const SKIP_HEADERS = ['日付', '曜日', '日付/曜日', '日付・曜日', '開館時間', '日', '曜', '備考', '特記事項',
  '不足', '合計', 'メモ', 'note', '氏名', '名前', 'スタッフ', '文', '理', '時間', '必要人数'];
const STAFF_HOLD_HOURS = 9;   // 職員の拘束時間（勤務開始からこの時間ぶんが出勤）
const SKIP_HEADER_RE = /日付|曜日|特記|備考|不足|合計|メモ|人数|区分|氏名|名前|スタッフ|時間$/;
const WD = ['日', '月', '火', '水', '木', '金', '土'];
const WD_RE = /^[月火水木金土日]$/;

/* ---------- 小道具 ---------- */
function cellText(v) {
  if (v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return v.getUTCFullYear() + '-' + String(v.getUTCMonth() + 1).padStart(2, '0') + '-' + String(v.getUTCDate()).padStart(2, '0');
  }
  if (typeof v === 'object') {
    if (v.richText) return v.richText.map(t => t.text).join('');
    if (v.result !== undefined && v.result !== null) return cellText(v.result);
    if (v.text !== undefined) return String(v.text);
    return '';
  }
  return String(v);
}

/** 全角の数字・記号を半角にそろえる */
function normalizeText(v) {
  return String(v == null ? '' : v)
    .replace(/[０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0))
    .replace(/[－ー―‐–—〜～~]/g, '-')
    .replace(/[．]/g, '.').replace(/[：]/g, ':').replace(/[／]/g, '/')
    .replace(/[（]/g, '(').replace(/[）]/g, ')');
}

function toMinutes(t) {
  t = String(t).trim();
  if (t.indexOf(':') >= 0) { const p = t.split(':'); return parseInt(p[0], 10) * 60 + parseInt(p[1], 10); }
  return Math.round(parseFloat(t) * 60);
}

/** セル内で折り返された文字をつなぐ。時間帯が途中で切れている場合はそのまま連結する */
function joinCellText(a, b) {
  if (!a) return b || '';
  if (!b) return a;
  if (/-\s*$/.test(a) || /^\s*-/.test(b)) return a + b;
  return a + '、' + b;
}

/** セル文字列から勤務時間帯 [[開始分,終了分],...] */
function parseRanges(text, open, close) {
  if (text === null || text === undefined) return [];
  const s = normalizeText(text).trim();
  if (!s) return [];
  if (['×', 'x', 'X', '✕', '-', '−', '休', '休み', 'なし'].indexOf(s) >= 0) return [];
  if (s.indexOf('終日') >= 0 || s.indexOf('通し') >= 0) return (open == null) ? [] : [[open, close]];
  const out = [];
  s.split(/[\n,、／\/]+/).forEach(seg => {
    const m = seg.match(/(\d{1,2}(?::\d{2}|\.\d+)?)\s*-\s*(\d{1,2}(?::\d{2}|\.\d+)?)/);
    if (!m) return;
    const a = toMinutes(m[1]), b = toMinutes(m[2]);
    if (b > a && a >= 0 && b <= 24 * 60) out.push([a, b]);
  });
  return out.sort((p, q) => p[0] - q[0]);
}

/** 「12」「8.5」のように数字だけの記入を勤務開始時刻として読む */
function parseStartOnly(text) {
  const s = normalizeText(text).trim();
  const m = s.match(/^(\d{1,2}(?:\.\d+)?)$/);
  if (!m) return null;
  const v = Math.round(parseFloat(m[1]) * 60);
  return (v >= 0 && v < 24 * 60) ? v : null;
}

/** 開館時間からはみ出す部分を落とす */
function clipRanges(ranges, open, close) {
  if (open == null || close == null) return ranges;
  const out = [];
  ranges.forEach(r => {
    const a = Math.max(r[0], open), b = Math.min(r[1], close);
    if (b > a) out.push([a, b]);
  });
  return out;
}

/** セル1つを勤務時間帯に変換する。
 *  hours を渡すと、数字だけの記入を「開始時刻＋hours時間」として扱う（職員用）。 */
function parseCell(text, open, close, hours, clip) {
  let rg = parseRanges(text, open, close);
  if (!rg.length && hours) {
    const st = parseStartOnly(text);
    if (st != null) rg = [[st, st + Math.round(hours * 60)]];
  }
  return clip === false ? rg : clipRanges(rg, open, close);
}

function isWorking(ranges, idx) {
  const s = DAY_START + idx * SLOT, e = s + SLOT;
  return ranges.some(r => r[0] < e && r[1] > s);
}

/** 日付らしさの判定。'full'=年月日 / 'md'=月日 / 'day'=日のみ */
function readDate(text) {
  const t = normalizeText(text).trim();
  if (!t) return null;
  let m = t.match(/^(\d{4})\s*[\/\-年.]\s*(\d{1,2})\s*[\/\-月.]\s*(\d{1,2})/);
  if (m) return { mode: 'full', y: +m[1], m: +m[2], d: +m[3] };
  m = t.match(/^(\d{1,2})\s*[\/\-月.]\s*(\d{1,2})\s*日?\s*[(（]?/);
  if (m && +m[1] >= 1 && +m[1] <= 12 && +m[2] >= 1 && +m[2] <= 31) return { mode: 'md', m: +m[1], d: +m[2] };
  m = t.match(/^(\d{1,2})\s*日?\s*[(（]?[月火水木金土日]?[)）]?$/);
  if (m && +m[1] >= 1 && +m[1] <= 31) return { mode: 'day', d: +m[1] };
  return null;
}

/** 「2026年9月」「9月」などから年月を拾う */
function detectYearMonth(allText) {
  let m = allText.match(/(20\d{2})\s*年\s*(\d{1,2})\s*月/);
  if (m) return { y: +m[1], m: +m[2] };
  m = allText.match(/(20\d{2})[\/\-](\d{1,2})(?![\/\-]\d)/);
  if (m) return { y: +m[1], m: +m[2] };
  m = allText.match(/(\d{1,2})\s*月/);
  if (m) return { y: null, m: +m[1] };
  return null;
}

const transpose = g => {
  const w = g.reduce((a, r) => Math.max(a, r.length), 0);
  return Array.from({ length: w }, (_, c) => g.map(r => r[c] || ''));
};

/* ---------- Excel → グリッド ---------- */
function gridFromWorksheet(ws) {
  const rows = [];
  const maxC = Math.min(ws.columnCount || 40, 80);
  for (let r = 1; r <= ws.rowCount; r++) {
    const row = ws.getRow(r), out = [];
    for (let c = 1; c <= maxC; c++) out.push(cellText(row.getCell(c).value).trim());
    rows.push(out);
  }
  return rows;
}

/* ---------- PDF → グリッド ---------- */
/** 1ページ分のテキスト片を、行×列のグリッドに組み直す */
function gridFromItems(items) {
  const src = items.filter(it => String(it.s).trim() !== '');
  if (!src.length) return { grid: [], ys: [] };
  const hs = src.map(it => it.h).sort((a, b) => a - b);
  const H = hs[Math.floor(hs.length / 2)] || 10;

  // 行：y座標が近いものをまとめる
  const sorted = src.slice().sort((a, b) => (b.y - a.y) || (a.x - b.x));
  const lines = [];
  sorted.forEach(it => {
    const last = lines[lines.length - 1];
    if (last && Math.abs(last.y - it.y) <= H * 0.6) { last.items.push(it); last.y = (last.y + it.y) / 2; }
    else lines.push({ y: it.y, items: [it] });
  });

  // 列：セルの左端（または中央）がそろう位置を列の基準にする。
  //     セル内の文字が隣の空きセルにはみ出しているPDFでも列を取り違えないため。
  const rowsCount = lines.length;
  const clusterBy = pick => {
    const pts = [];
    lines.forEach((line, li) => line.items.forEach(it => pts.push({ v: pick(it), li: li })));
    pts.sort((a, b) => a.v - b.v);
    const cl = [];
    pts.forEach(pt => {
      const last = cl[cl.length - 1];
      if (last && pt.v - last.v <= H * 0.6) { last.v = (last.v * last.n + pt.v) / (last.n + 1); last.n++; last.rows.add(pt.li); }
      else cl.push({ v: pt.v, n: 1, rows: new Set([pt.li]) });
    });
    const need = Math.max(2, Math.round(rowsCount * 0.15));
    return cl.filter(c => c.rows.size >= need).map(c => c.v);
  };
  const byLeft = clusterBy(it => it.x);
  const byCenter = clusterBy(it => it.x + Math.max(it.w, 1) / 2);
  let anchors = (byLeft.length >= byCenter.length ? byLeft : byCenter);
  let pickPos = (byLeft.length >= byCenter.length) ? (it => it.x) : (it => it.x + Math.max(it.w, 1) / 2);

  let colOf;
  const usingAnchors = anchors.length >= 3;
  if (usingAnchors) {
    colOf = it => {
      const v = pickPos(it);
      let best = 0, bd = Infinity;
      anchors.forEach((a, i) => { const d = Math.abs(v - a); if (d < bd) { bd = d; best = i; } });
      return best;
    };
  } else {
    // そろった位置が見つからないときは、文字の占める区間の切れ目で列を切る
    const spans = src.map(it => [it.x, it.x + Math.max(it.w, 1)]).sort((a, b) => a[0] - b[0]);
    const blocks = [];
    spans.forEach(sp => {
      const last = blocks[blocks.length - 1];
      if (last && sp[0] <= last[1] + H * 0.4) last[1] = Math.max(last[1], sp[1]);
      else blocks.push([sp[0], sp[1]]);
    });
    anchors = blocks;
    colOf = it => {
      const c = it.x + Math.max(it.w, 1) / 2;
      let best = 0, bd = Infinity;
      blocks.forEach((b, i) => {
        const d = (c >= b[0] && c <= b[1]) ? 0 : Math.min(Math.abs(c - b[0]), Math.abs(c - b[1]));
        if (d < bd) { bd = d; best = i; }
      });
      return best;
    };
  }
  const colCount = anchors.length;

  const rows = [], ys = [];
  lines.forEach(line => {
    // まず各文字を列に割り当てる（左から順に見て、すき間なく続く文字は同じセルの続きとみなす）
    const bucket = [];
    let prevRight = -1e9, prevCol = 0, prevY = null;
    line.items.slice().sort((a, b) => a.x - b.x).forEach(it => {
      const own = colOf(it);
      const onAnchor = usingAnchors && Math.abs(pickPos(it) - anchors[own]) <= H * 0.35;
      const sameLevel = prevY !== null && Math.abs(it.y - prevY) <= H * 0.3;
      const cont = !onAnchor && sameLevel && it.x <= prevRight + H * 0.15;
      const i = cont ? prevCol : own;
      (bucket[i] = bucket[i] || []).push(it);
      prevRight = it.x + Math.max(it.w, 1); prevCol = i; prevY = it.y;
    });
    // セルの中身を組み立てる。セル内で上下2段になっている場合は上の段から読む
    const cells = [];
    for (let i = 0; i < colCount; i++) {
      const list = bucket[i];
      if (!list) { cells[i] = ''; continue; }
      list.sort((a, b) => (b.y - a.y) || (a.x - b.x));
      let text = '', lastY = null;
      list.forEach(it => {
        if (lastY === null) text = it.s;
        else if (Math.abs(it.y - lastY) <= H * 0.3) text += it.s;   // 同じ段の続き
        else text = joinCellText(text, it.s);                       // 折り返した次の段
        lastY = it.y;
      });
      cells[i] = text.trim();
    }
    rows.push(cells);
    ys.push(line.y);
  });
  return { grid: rows, ys: ys };
}

/* ---------- グリッド → シフト ---------- */
function scoreDateColumn(grid, c) {
  const entries = [];
  let mdCount = 0;
  for (let r = 0; r < grid.length; r++) {
    const d = readDate(grid[r][c]);
    if (!d) continue;
    if (d.mode !== 'day') mdCount++;
    entries.push({ r: r, d: d });
  }
  if (entries.length < 3) return { score: 0 };
  // 「日のみ」の列は、1〜31が増えていく並びのときだけ日付とみなす
  if (!mdCount) {
    let ok = true;
    for (let i = 1; i < entries.length; i++) if (entries[i].d.d <= entries[i - 1].d.d) { ok = false; break; }
    if (!ok || entries.length < 10) return { score: 0 };
  }
  return { score: entries.length + mdCount * 0.5, entries: entries, col: c };
}

function analyzeGrid(grid, fallbackYM, ys, keepNames) {
  const keep = keepNames || [];
  const width = grid.reduce((a, r) => Math.max(a, r.length), 0);
  let best = { score: 0 };
  for (let c = 0; c < width; c++) {
    const s = scoreDateColumn(grid, c);
    if (s.score > best.score) best = s;
  }
  if (!best.score) return null;

  const dateCol = best.col, firstRow = best.entries[0].r;

  // 見出し行：「日付」がある行、なければ日付行の直前で中身のある行
  let headerRow = -1;
  for (let r = 0; r < grid.length; r++) {
    if (String(grid[r][dateCol] || '').trim() === '日付') { headerRow = r; break; }
  }
  if (headerRow < 0) {
    for (let r = firstRow - 1; r >= 0; r--) {
      if (grid[r].some((t, i) => i !== dateCol && String(t).trim() !== '')) { headerRow = r; break; }
    }
  }
  if (headerRow < 0) headerRow = Math.max(0, firstRow - 1);

  const dateRows = best.entries.map(e => e.r);
  const isWeekdayCol = c => {
    let n = 0, w = 0;
    dateRows.forEach(r => { const t = String(grid[r][c] || '').trim(); if (t) { n++; if (WD_RE.test(t)) w++; } });
    return n >= 3 && w / n >= 0.6;
  };

  const staff = [];
  let openCol = -1;
  for (let c = 0; c < width; c++) {
    if (c === dateCol) continue;
    const h = String(grid[headerRow][c] || '').trim();
    if (h === '開館時間') { openCol = c; continue; }
    if (!h || SKIP_HEADERS.indexOf(h) >= 0 || SKIP_HEADER_RE.test(h) || readDate(h) || WD_RE.test(h)) continue;
    if (isWeekdayCol(c)) continue;
    if (staff.some(s => s.name === h)) continue;
    // 時間帯が1つも入っていない列はスタッフ列とみなさない（タイトルの誤検出よけ）。
    // ただし職員として指定された名前の列は、数字だけの記入でも残す。
    if (keep.indexOf(h) < 0) {
      const hasTime = dateRows.some(r => parseRanges(grid[r][c]).length > 0 || String(grid[r][c] || '').indexOf('終日') >= 0);
      if (!hasTime) continue;
    }
    staff.push({ name: h, col: c });
  }
  if (!staff.length) return null;

  const allText = grid.map(r => r.join(' ')).join(' ');
  const ym = detectYearMonth(allText) || {};
  const year = ym.y || (fallbackYM && fallbackYM.y) || new Date().getFullYear();
  const month = ym.m || (fallbackYM && fallbackYM.m) || (new Date().getMonth() + 1);
  let guessed = false;

  const days = best.entries.map((e, i) => {
    const dd = e.d;
    let y = dd.y, m = dd.m, d = dd.d;
    if (dd.mode !== 'full') { guessed = true; if (m === undefined) m = month; y = year; }
    let open = null, close = null;
    if (openCol >= 0) {
      const oc = parseRanges(grid[e.r][openCol]);
      if (oc.length) { open = oc[0][0]; close = oc[0][1]; }
    }
    const cells = {};
    const nextRow = (best.entries[i + 1] ? best.entries[i + 1].r : grid.length);
    const stop = Math.min(nextRow, e.r + 4);
    staff.forEach(s => {
      let t = String(grid[e.r][s.col] || '').trim();
      for (let rr = e.r + 1; rr < stop; rr++) {
        const x = String(grid[rr][s.col] || '').trim();
        if (x) t = joinCellText(t, x);
      }
      if (t) cells[s.name] = t;
    });
    return { y: y, m: m, d: d, wd: new Date(Date.UTC(y, m - 1, d)).getUTCDay(), open: open, close: close, cells: cells, y0: ys ? ys[e.r] : null };
  });

  const filled = days.filter(day => Object.keys(day.cells).some(n => parseCell(day.cells[n], day.open, day.close, 9).length));
  return {
    score: best.score + filled.length,
    staffNames: staff.map(s => s.name),
    days: days,
    guessedMonth: guessed,
    filledDays: filled.length
  };
}

/** グリッド（縦横どちらでも可）を解析する */
function parseGrid(grid, fallbackYM, ys, keepNames) {
  const a = analyzeGrid(grid, fallbackYM, ys, keepNames);
  const b = analyzeGrid(transpose(grid), fallbackYM, null, keepNames);
  if (!a && !b) return null;
  if (!a) return b;
  if (!b) return a;
  return (b.score > a.score) ? b : a;
}

/** 複数ページ・複数シートの結果をまとめる */
function mergeParsed(list) {
  const byDate = {}, order = [], names = [];
  let guessed = false;
  list.forEach(p => {
    if (!p) return;
    if (p.guessedMonth) guessed = true;
    p.staffNames.forEach(n => { if (names.indexOf(n) < 0) names.push(n); });
    p.days.forEach(day => {
      const k = day.y + '-' + day.m + '-' + day.d;
      if (!byDate[k]) { byDate[k] = { y: day.y, m: day.m, d: day.d, wd: day.wd, open: day.open, close: day.close, cells: {} }; order.push(k); }
      const t = byDate[k];
      if (day.open != null) { t.open = day.open; t.close = day.close; }
      Object.keys(day.cells).forEach(n => { if (!t.cells[n]) t.cells[n] = day.cells[n]; });
    });
  });
  if (!order.length) return null;
  const days = order.map(k => byDate[k]).sort((x, y) => (x.y - y.y) || (x.m - y.m) || (x.d - y.d));
  return { staffNames: names, days: days, guessedMonth: guessed };
}

/** セル内で折り返された行を、上の行に取り込む
 *  （左端の見出し列が空で、文字の入っている列が上の行の一部に収まっている行が折り返し） */
function mergeWrappedRows(page) {
  const grid = [], ys = [];
  page.grid.forEach((row, i) => {
    const prev = grid[grid.length - 1];
    const cols = row.map((t, c) => (t ? c : -1)).filter(c => c >= 0);
    const isWrap = prev && cols.length > 0 && !row[0] &&
      !row.some(t => readDate(t)) &&
      cols.every(c => prev[c] && prev[c] !== '');
    if (isWrap) {
      cols.forEach(c => { prev[c] = joinCellText(prev[c], row[c]); });
    } else { grid.push(row.slice()); ys.push(page.ys[i]); }
  });
  return { grid: grid, ys: ys };
}

/** 複数ページのPDFを、行のy座標をたよりに1枚の表につなぎ直す
 *  （表が横に分割されて印刷されている場合は横につなぎ、そうでなければ下に足す） */
function stitchPages(pages) {
  if (!pages.length) return { grid: [], ys: [] };
  const base = { grid: pages[0].grid.map(r => r.slice()), ys: pages[0].ys.slice() };
  for (let i = 1; i < pages.length; i++) {
    const pg = pages[i];
    if (!pg.ys.length) continue;
    let match = 0;
    pg.ys.forEach(y => { if (base.ys.some(by => Math.abs(by - y) <= 3)) match++; });
    const width = base.grid.reduce((a, r) => Math.max(a, r.length), 0);
    if (match / pg.ys.length >= 0.6) {
      pg.grid.forEach((row, ri) => {
        let bi = -1, bd = 3.5;
        base.ys.forEach((by, k) => { const d = Math.abs(by - pg.ys[ri]); if (d < bd) { bd = d; bi = k; } });
        if (bi < 0) {
          const nr = [];
          for (let c = 0; c < width; c++) nr.push('');
          row.forEach((t, c) => { nr[width + c] = t; });
          base.grid.push(nr); base.ys.push(pg.ys[ri]);
        } else {
          const r2 = base.grid[bi];
          while (r2.length < width) r2.push('');
          row.forEach((t, c) => { r2[width + c] = t; });
        }
      });
    } else {
      pg.grid.forEach((row, ri) => { base.grid.push(row.slice()); base.ys.push(pg.ys[ri]); });
    }
  }
  const W = base.grid.reduce((a, r) => Math.max(a, r.length), 0);
  base.grid.forEach(r => { for (let c = 0; c < W; c++) if (r[c] === undefined) r[c] = ''; });
  return base;
}


/* ============================ アプリ本体 ============================ */
const CAP_DEFAULT = 6, TARGET_DEFAULT = 3;
const DEFAULT_TEAMS = {
  'ふか': ['深見', '久保田', '梶田'],
  'わた': ['渡邊', '廣中', '増田', '遅野井'],
  'むら': ['村山', '西村', '平川'],
  'てつ': ['手塚', '岡田', '森'],
  'うだ': ['岩本', '宇田川', '松岡']
};
const TEAM_COLOR = { 'ふか': '#4E8FC7', 'わた': '#D98A45', 'むら': '#5FA96E', 'てつ': '#8C7BC0', 'うだ': '#C9A227', '全体': '#8A8A8A' };
const WDAY = ['日', '月', '火', '水', '木', '金', '土'];
const KEY = 'meeting-booking-v1';

const state = {
  tab: 'book',
  month: null, year: null,
  students: [],          // {name,kana,grade,team,skip,marks:{day:mark}}
  days: [],              // [1..31]
  shiftDays: [],         // [{m,d,staff:[名前]}]
  pending: [],           // {name,team,day,at,no,id}  no=マナビス生番号 / id=同期先での番号
  noByName: {},          // 氏名 → 生番号
  sync: { url: '', token: '' },
  syncAt: null, syncErr: '',
  loadedAt: null, shiftAt: null,
  teams: JSON.parse(JSON.stringify(DEFAULT_TEAMS)),
  cap: CAP_DEFAULT, target: TARGET_DEFAULT,
  q: '', filterTeam: '', onlyNeed: true,
  open: null             // 開いている生徒
};

const $ = id => document.getElementById(id);
const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const pad = n => String(n).padStart(2, '0');

/* ---------- 保存（Safariで開いたときも、Claude内でもどちらでも動く） ---------- */
const store = {
  async get(k) {
    try { if (window.storage) { const r = await window.storage.get(k); return r ? r.value : null; } } catch (e) { }
    try { return window.localStorage.getItem(k); } catch (e) { return null; }
  },
  async set(k, v) {
    try { if (window.storage) { await window.storage.set(k, v); return; } } catch (e) { }
    try { window.localStorage.setItem(k, v); } catch (e) { }
  }
};
async function save() {
  await store.set(KEY, JSON.stringify({
    month: state.month, year: state.year, students: state.students, days: state.days,
    shiftDays: state.shiftDays, pending: state.pending, loadedAt: state.loadedAt,
    shiftAt: state.shiftAt, teams: state.teams, cap: state.cap, target: state.target,
    noByName: state.noByName, sync: state.sync, syncAt: state.syncAt
  }));
}
async function load() {
  const raw = await store.get(KEY);
  if (!raw) return;
  try { Object.assign(state, JSON.parse(raw)); } catch (e) { }
}

/* ---------- Excelの読み取り ---------- */
function cellStr(cell) {
  if (!cell) return '';
  const v = cell.v;
  if (Object.prototype.toString.call(v) === '[object Date]') return v.getFullYear() + '-' + pad(v.getMonth() + 1) + '-' + pad(v.getDate());
  return String(v === null || v === undefined ? '' : v).trim();
}
function gridFrom(ws, maxRows, maxCols) {
  if (!ws || !ws['!ref']) return [];
  const rg = XLSX.utils.decode_range(ws['!ref']);
  const rows = [];
  for (let r = rg.s.r; r <= Math.min(rg.e.r, rg.s.r + (maxRows || 2000)); r++) {
    const row = [];
    for (let c = rg.s.c; c <= Math.min(rg.e.c, maxCols || 80); c++) {
      row.push(cellStr(ws[XLSX.utils.encode_cell({ r: r, c: c })]));
    }
    rows.push(row);
  }
  return rows;
}

/** 氏名の表記ゆれを吸収する（全角・半角スペースを取る） */
function squashName(s) {
  return String(s == null ? '' : s).replace(/[\s\u3000]/g, '').trim();
}

/** 「顧客情報貼り付け」シートから 氏名→マナビス生番号 の対応を作る */
function readNoByName(wb) {
  const map = {};
  const ws = wb.Sheets['顧客情報貼り付け'];
  if (!ws) return map;
  const g = gridFrom(ws, 500, 12);
  if (!g.length) return map;
  let noCol = -1, nameCol = -1;
  for (let c = 0; c < g[0].length; c++) {
    const h = String(g[0][c] || '').trim();
    if (h.indexOf('生番号') >= 0) noCol = c;
    else if (h === '氏名' && nameCol < 0) nameCol = c;
  }
  if (noCol < 0 || nameCol < 0) return map;
  for (let r = 1; r < g.length; r++) {
    const no = String(g[r][noCol] || '').trim();
    const nm = squashName(g[r][nameCol]);
    if (/^\d{4,8}$/.test(no) && nm) map[nm] = no;
  }
  return map;
}

/** 日程管理シート（mainシート）を読む */
function parseSchedule(wb) {
  const ws = wb.Sheets['main'] || wb.Sheets[wb.SheetNames[0]];
  const g = gridFrom(ws, 400, 45);
  const at = (r, c) => (g[r - 1] && g[r - 1][c - 1]) || '';
  const month = Number(at(2, 1));
  if (!month || month < 1 || month > 12) throw new Error('A2セルから月を読み取れませんでした。日程管理シートのファイルか確認してください。');
  const dayCols = {};
  for (let c = 1; c <= 45; c++) {
    const n = Number(at(8, c));
    if (Number.isInteger(n) && n >= 1 && n <= 31 && !dayCols[n]) dayCols[n] = c;
  }
  const days = Object.keys(dayCols).map(Number).sort((a, b) => a - b);
  if (!days.length) throw new Error('8行目から日付を読み取れませんでした。');
  const students = [];
  for (let r = 10; r <= g.length; r++) {
    const name = at(r, 1);
    if (!name) continue;
    const marks = {};
    days.forEach(d => { const t = at(r, dayCols[d]); if (t) marks[d] = t; });
    students.push({
      name: name, kana: at(r, 2), grade: normalizeText(at(r, 3)),
      team: at(r, 4) || '未設定', skip: at(r, 5), marks: marks
    });
  }
  if (!students.length) throw new Error('生徒の行が見つかりませんでした。');
  return { month: month, days: days, students: students };
}

/** シフト表（Excel / PDF）を読んで、日ごとの出勤者を出す */
async function parseShiftFile(file) {
  let parsed;
  if (/\.pdf$/i.test(file.name)) {
    const lib = window.pdfjsLib || window['pdfjs-dist/build/pdf'];
    if (!lib) throw new Error('PDFの読み取り機能を読み込めませんでした。');
    try { lib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js'; } catch (e) { }
    const doc = await lib.getDocument({ data: new Uint8Array(await file.arrayBuffer()) }).promise;
    const pages = [];
    for (let p = 1; p <= doc.numPages; p++) {
      const page = await doc.getPage(p);
      const tc = await page.getTextContent();
      const items = tc.items.filter(i => i.str && i.str.trim()).map(i => ({
        s: i.str, x: i.transform[4], y: i.transform[5], w: i.width || 0, h: Math.abs(i.transform[3]) || 10
      }));
      if (items.length) pages.push(gridFromItems(items));
    }
    if (!pages.length) throw new Error('PDFから文字を読み取れませんでした。');
    const st = mergeWrappedRows(stitchPages(pages));
    parsed = mergeParsed([parseGrid(st.grid, null, st.ys)]);
  } else {
    const wb = XLSX.read(new Uint8Array(await file.arrayBuffer()), { cellDates: true, type: 'array' });
    parsed = mergeParsed(wb.SheetNames.map(n => {
      try { return parseGrid(gridFrom(wb.Sheets[n], 400, 60)); } catch (e) { return null; }
    }));
  }
  if (!parsed) throw new Error('シフト表から日付とスタッフ名を読み取れませんでした。');
  return parsed.days.map(d => ({
    y: d.y, m: d.m, d: d.d,
    staff: Object.keys(d.cells).filter(n => parseCell(d.cells[n], d.open, d.close).length)
  }));
}

/* ---------- 空き状況の計算 ---------- */
function teamOf(staffName) {
  for (const t in state.teams) if (state.teams[t].indexOf(staffName) >= 0) return t;
  return null;
}
function staffOn(day) {
  const rec = state.shiftDays.find(x => x.m === state.month && x.d === day);
  return rec ? rec.staff : null;      // null＝シフト未読込
}
function capacityOf(day, team) {
  const staff = staffOn(day);
  if (staff === null) return { staff: [], cap: null };
  const list = (team === '全体' || !state.teams[team])
    ? staff.filter(n => teamOf(n))
    : staff.filter(n => teamOf(n) === team);
  return { staff: list, cap: list.length * state.cap };
}
function loadOf(day, team) {
  let n = 0;
  state.students.forEach(s => {
    if (s.team !== team) return;
    const m = s.marks[day];
    if (m === '予定日' || m === '済') n++;
  });
  state.pending.forEach(p => { if (p.day === day && p.team === team) n++; });
  return n;
}
function doneCount(s) {
  let n = Object.keys(s.marks).filter(d => s.marks[d] === '予定日' || s.marks[d] === '済').length;
  n += state.pending.filter(p => p.name === s.name).length;
  return n;
}
function datesOf(s) {
  const ds = Object.keys(s.marks).filter(d => s.marks[d] === '予定日' || s.marks[d] === '済').map(Number);
  state.pending.filter(p => p.name === s.name).forEach(p => ds.push(p.day));
  return ds.sort((a, b) => a - b);
}
function needsRebook(s) {
  return Object.keys(s.marks).some(d => s.marks[d] === '要振替');
}

/* ---------- 画面 ---------- */
/* ---------- スプレッドシートとの同期 ---------- */
function syncReady() { return !!(state.sync && state.sync.url && state.sync.token); }

async function syncCall(action, extra) {
  const body = Object.assign({ token: state.sync.token, action: action }, extra || {});
  const res = await fetch(state.sync.url, {
    method: 'POST',
    // text/plain にしておくと、事前確認の通信が起きず確実につながる
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify(body)
  });
  const text = await res.text();
  if (text.trim().charAt(0) === '<') {
    // JSONではなくHTMLが返ってきている＝Googleのログイン画面
    throw new Error('公開設定を見直してください（デプロイを管理 → アクセスできるユーザーを「全員」に）');
  }
  let json;
  try { json = JSON.parse(text); }
  catch (e) { throw new Error('応答を読み取れません。URLが正しいか確認してください'); }
  if (!json.ok) throw new Error(json.error || '同期に失敗しました');
  return json;
}

/** 同期先の一覧を、手元の未反映一覧に取り込む */
function mergeRemote(items) {
  const nameByNo = {};
  Object.keys(state.noByName).forEach(nm => { nameByNo[state.noByName[nm]] = nm; });
  const byName = {};
  state.students.forEach(s => { byName[squashName(s.name)] = s; });

  // 同期先にある分（＝送信済み）を作り直す
  const remote = items.map(it => {
    const nm = nameByNo[String(it.no)] || '';
    const st = nm ? byName[squashName(nm)] : null;
    return {
      id: it.id, no: String(it.no), day: Number(it.d),
      name: nm || ('生番号 ' + it.no),
      team: (st && st.team) || it.team || ''
    };
  });
  // まだ送れていない分は手元に残す
  const unsent = state.pending.filter(p => !p.id);
  const seen = {};
  remote.forEach(r => { seen[r.no + '/' + r.day] = true; });
  state.pending = remote.concat(unsent.filter(u => !seen[u.no + '/' + u.day]));
}

/** 未送信の分を送って、最新の一覧を受け取る */
async function syncPush(silent) {
  if (!syncReady()) return;
  const unsent = state.pending.filter(p => !p.id && p.no);
  try {
    const r = unsent.length ? await syncCall('add', {
      items: unsent.map(p => ({ m: state.month, d: p.day, no: p.no, team: p.team }))
    }) : await syncCall('list', {});
    mergeRemote(r.items || []);
    state.syncAt = Date.now(); state.syncErr = '';
    await save();
    if (!silent) toast(unsent.length ? '同期しました' : '最新の状態です');
  } catch (e) {
    state.syncErr = e.message;
    if (!silent) toast('同期できません：' + e.message);
  }
  renderHeader();
}

async function syncPull(silent) { return syncPush(silent); }

function toast(msg) {
  const t = $('toast'); t.textContent = msg; t.classList.add('show');
  clearTimeout(toast._t); toast._t = setTimeout(() => t.classList.remove('show'), 2000);
}

function renderHeader() {
  $('hdMon').textContent = state.month ? state.month + '月' : '未読込';
  const n = state.pending.length;
  const unsent = syncReady() ? state.pending.filter(p => !p.id).length : 0;
  $('hdBadge').style.display = n ? 'inline-block' : 'none';
  $('hdBadge').textContent = unsent ? '未送信 ' + unsent : '未反映 ' + n;
  $('hdBadge').style.background = unsent ? '#B23A2E' : '';
  $('hdBadge').style.color = unsent ? '#fff' : '';
  const s = $('stale');
  if (state.loadedAt) {
    const days = Math.floor((Date.now() - state.loadedAt) / 86400000);
    if (days >= 3) { s.style.display = 'block'; s.textContent = '前回の読み込みから' + days + '日たっています。Excelを読み込み直してください。'; }
    else s.style.display = 'none';
  } else s.style.display = 'none';
}

const ICONS = {
  book: '<path d="M12 5v14M5 12h14"/>',
  days: '<rect x="3" y="5" width="18" height="16" rx="2"/><path d="M3 10h18M8 3v4M16 3v4"/>',
  need: '<path d="M12 3l9 16H3z"/><path d="M12 10v4M12 17h.01"/>',
  data: '<path d="M4 7v10c0 1.7 3.6 3 8 3s8-1.3 8-3V7"/><ellipse cx="12" cy="7" rx="8" ry="3"/>'
};
function renderNav() {
  const tabs = [['book', '設定する'], ['days', '日別の空き'], ['need', '未設定'], ['data', 'データ']];
  $('nav').innerHTML = tabs.map(([k, label]) =>
    `<button data-tab="${k}" class="${state.tab === k ? 'on' : ''}">
      <svg viewBox="0 0 24 24">${ICONS[k]}</svg>${label}
      ${k === 'data' && state.pending.length ? '<span class="dot"></span>' : ''}</button>`).join('');
  $('nav').querySelectorAll('button').forEach(b =>
    b.addEventListener('click', () => { state.tab = b.dataset.tab; render(); autoSync(); }));
}

let lastAutoSync = 0;
function autoSync() {
  if (!syncReady()) return;
  if (Date.now() - lastAutoSync < 15000) return;   // 連打しない
  lastAutoSync = Date.now();
  syncPush(true);
}

function render() {
  renderHeader(); renderNav();
  const v = $('view');
  if (!state.students.length && state.tab !== 'data') {
    v.innerHTML = `<div class="empty"><b>まだデータがありません</b>
      「データ」から日程管理シートとシフト表を読み込んでください。</div>`;
    return;
  }
  if (state.tab === 'book') renderBook(v);
  else if (state.tab === 'days') renderDays(v);
  else if (state.tab === 'need') renderNeed(v);
  else renderData(v);
}

/* --- 生徒を選んで設定する --- */
function renderBook(v) {
  const teams = Object.keys(state.teams).concat(['全体']);
  v.innerHTML = `
    <div class="search"><input type="text" id="q" placeholder="名前・カナで検索" value="${esc(state.q)}"></div>
    <div class="chips">
      <button class="chip ${state.onlyNeed ? 'on' : ''}" id="cNeed">未設定のみ</button>
      <button class="chip ${!state.filterTeam ? 'on' : ''}" data-team="">全チーム</button>
      ${teams.map(t => `<button class="chip ${state.filterTeam === t ? 'on' : ''}" data-team="${t}">${t}</button>`).join('')}
    </div>
    <div id="list"></div>`;
  $('q').addEventListener('input', e => { state.q = e.target.value; renderList(); });
  $('cNeed').addEventListener('click', () => { state.onlyNeed = !state.onlyNeed; render(); });
  v.querySelectorAll('[data-team]').forEach(b =>
    b.addEventListener('click', () => { state.filterTeam = b.dataset.team; render(); }));
  renderList();
}

function matchedStudents() {
  const q = state.q.trim();
  return state.students.filter(s => {
    if (s.skip) return false;
    if (state.filterTeam && s.team !== state.filterTeam) return false;
    if (q && s.name.indexOf(q) < 0 && (s.kana || '').indexOf(q) < 0) return false;
    if (state.onlyNeed && doneCount(s) >= state.target && !needsRebook(s)) return false;
    return true;
  });
}

function renderList() {
  const list = matchedStudents();
  const box = $('list');
  if (!list.length) { box.innerHTML = `<div class="empty">該当する生徒がいません</div>`; return; }
  box.innerHTML = list.slice(0, 120).map((s, i) => {
    const done = doneCount(s), rest = Math.max(0, state.target - done);
    const ds = datesOf(s);
    const cls = rest === 0 ? 'done' : (rest >= state.target ? 'need' : '');
    return `<button class="stu ${cls}" data-i="${state.students.indexOf(s)}" style="--tc:${TEAM_COLOR[s.team] || '#CCC'}">
      <span class="nm"><b>${esc(s.name)}</b>
        <small>${esc(s.grade)} ・ ${esc(s.team)}${ds.length ? ' ・ ' + ds.map(d => state.month + '/' + d).join(' ') : ''}</small></span>
      <span class="rest"><b>${rest}</b><small>${needsRebook(s) ? '要振替' : '残り'}</small></span>
    </button>`;
  }).join('') + (list.length > 120 ? `<div class="empty">ほか${list.length - 120}名。検索で絞り込んでください</div>` : '');
  box.querySelectorAll('.stu').forEach(b =>
    b.addEventListener('click', () => openStudent(Number(b.dataset.i))));
}

/* --- 生徒ごとの日付選び --- */
function openStudent(idx) {
  state.open = idx;
  const s = state.students[idx];
  const wrap = document.createElement('div');
  wrap.className = 'sheet'; wrap.id = 'sheet';
  wrap.innerHTML = `<div class="sheet-in">
      <div class="sheet-hd">
        <span class="t"><b>${esc(s.name)}</b><small>${esc(s.grade)} ・ ${esc(s.team)}チーム ・ 残り${Math.max(0, state.target - doneCount(s))}回</small></span>
        <button class="x" id="closeSheet">閉じる</button>
      </div>
      <div class="setdates" id="setdates"></div>
      <div class="sheet-bd" id="dayList"></div>
    </div>`;
  // 統合版では #appB の中に入れる（外に出すとCSSが効かず、画面が重なって表示される）
  (document.getElementById('appB') || document.body).appendChild(wrap);
  wrap.addEventListener('click', e => { if (e.target === wrap) closeSheet(); });
  $('closeSheet').addEventListener('click', closeSheet);
  renderDayList();
}
function closeSheet() {
  const w = $('sheet'); if (w) w.remove();
  state.open = null; render();
}

function renderDayList() {
  const s = state.students[state.open];
  const mine = datesOf(s);
  $('setdates').innerHTML = mine.length
    ? '設定済み： ' + mine.map(d => `<span>${state.month}/${d}</span>`).join('')
    : '<span style="color:#5A6672">まだ1件も入っていません</span>';

  const today = new Date();
  const isThisMonth = (today.getMonth() + 1) === state.month;
  const html = state.days.map(d => {
    const wd = state.year ? new Date(state.year, state.month - 1, d).getDay() : null;
    const past = isThisMonth && d < today.getDate();
    const { staff, cap } = capacityOf(d, s.team);
    const used = loadOf(d, s.team);
    const rest = cap === null ? null : cap - used;
    const already = mine.indexOf(d) >= 0;
    let cls = 'r-ok', pct = 0;
    if (rest === null) { cls = 'r-warn'; }
    else { pct = cap ? Math.min(100, used / cap * 100) : 100; cls = rest <= 0 ? 'r-ng' : (rest <= 2 ? 'r-warn' : 'r-ok'); }
    const near = mine.some(x => Math.abs(x - d) > 0 && Math.abs(x - d) <= 2);
    const wdc = wd === 6 ? 'sat' : (wd === 0 ? 'sun' : '');
    return `<button class="day ${already ? 'mine' : ''} ${past ? 'full' : ''}" data-d="${d}" ${already ? 'disabled' : ''}>
        <span class="d ${wdc}"><b>${d}</b><small>${wd === null ? '' : WDAY[wd]}</small></span>
        <span class="info">
          <span class="bar"><i style="width:${pct}%;background:${cls === 'r-ng' ? '#B23A2E' : cls === 'r-warn' ? '#D99A2B' : '#3E9E7A'}"></i></span>
          <small>${cap === null ? 'シフト未読込' : (staff.length ? staff.join('・') : 'このチームの出勤なし')}</small>
        </span>
        <span class="rest ${cls}"><b>${already ? '設定済' : (rest === null ? '?' : rest)}</b>${already ? '' : '<small>あと</small>'}</span>
      </button>` + (near && !already ? `<div class="note">前後2日以内に予定があります</div>` : '');
  }).join('');
  $('dayList').innerHTML = html;
  $('dayList').querySelectorAll('.day').forEach(b => {
    if (b.disabled) return;
    b.addEventListener('click', () => confirmDay(Number(b.dataset.d)));
  });
}

async function confirmDay(day) {
  const s = state.students[state.open];
  const { cap } = capacityOf(day, s.team);
  const used = loadOf(day, s.team);
  if (cap !== null && cap - used <= 0) {
    const msg = cap === 0
      ? state.month + '/' + day + ' は' + s.team + 'チームの出勤がありません。それでも入れますか？'
      : state.month + '/' + day + ' は' + s.team + 'チームが上限（' + cap + '件）に達しています。それでも入れますか？';
    if (!confirm(msg)) return;
  }
  const no = state.noByName[squashName(s.name)] || '';
  state.pending.push({ name: s.name, team: s.team, day: day, at: Date.now(), no: no, id: null });
  await save();
  toast(state.month + '/' + day + ' に登録しました');
  renderDayList(); renderHeader();
  if (syncReady()) syncPush(true);   // 送信は裏で行う
}

/* --- 日別の空き --- */
function renderDays(v) {
  const teams = Object.keys(state.teams);
  const rows = state.days.map(d => {
    const wd = state.year ? new Date(state.year, state.month - 1, d).getDay() : null;
    const cells = teams.map(t => {
      const { cap } = capacityOf(d, t);
      const used = loadOf(d, t);
      if (cap === null) return `<td><span class="cell">-</span></td>`;
      const rest = cap - used;
      const bg = cap === 0 ? '#EFEDE6' : rest <= 0 ? 'var(--ng-bg)' : rest <= 2 ? 'var(--warn-bg)' : 'var(--ok-bg)';
      const fg = cap === 0 ? '#9AA0A6' : rest <= 0 ? 'var(--ng)' : rest <= 2 ? 'var(--warn)' : 'var(--ok)';
      return `<td><span class="cell" style="background:${bg};color:${fg}">${used}/${cap}</span></td>`;
    }).join('');
    const wc = wd === 6 ? 'sat' : (wd === 0 ? 'sun' : '');
    return `<tr><td class="dt ${wc}">${state.month}/${d}${wd === null ? '' : '(' + WDAY[wd] + ')'}</td>${cells}</tr>`;
  }).join('');
  v.innerHTML = `<div class="card" style="padding:0;overflow:auto;max-height:74vh">
      <table class="grid"><thead><tr><th style="text-align:left;padding-left:8px">日付</th>
      ${teams.map(t => `<th>${t}</th>`).join('')}</tr></thead><tbody>${rows}</tbody></table></div>
    <p style="font-size:12px;color:var(--ink-2);margin:0 4px">予定件数 / 受け入れ上限（そのチームの出勤人数 × ${state.cap}件）。灰色はそのチームの出勤がない日です。</p>`;
}

/* --- 未設定の生徒 --- */
function renderNeed(v) {
  const need = state.students.filter(s => !s.skip && (doneCount(s) < state.target || needsRebook(s)));
  const byTeam = {};
  need.forEach(s => { (byTeam[s.team] = byTeam[s.team] || []).push(s); });
  const total = need.length;
  v.innerHTML = `<div class="card"><h2>あと${total}名</h2>
      <p>目標${state.target}回に届いていない生徒と、要振替の生徒です。</p></div>` +
    Object.keys(byTeam).map(t => `<div class="card">
      <h2 style="color:${TEAM_COLOR[t] || '#333'}">${esc(t)}　${byTeam[t].length}名</h2>
      ${byTeam[t].map(s => {
        const rest = Math.max(0, state.target - doneCount(s));
        return `<button class="stu" data-i="${state.students.indexOf(s)}" style="--tc:${TEAM_COLOR[t] || '#CCC'}">
          <span class="nm"><b>${esc(s.name)}</b><small>${esc(s.grade)}${needsRebook(s) ? ' ・ 要振替' : ''}</small></span>
          <span class="rest"><b>${rest}</b><small>残り</small></span></button>`;
      }).join('')}</div>`).join('');
  v.querySelectorAll('.stu').forEach(b =>
    b.addEventListener('click', () => openStudent(Number(b.dataset.i))));
}

/* --- データ --- */
function renderData(v) {
  const lines = state.pending.slice().sort((a, b) => a.day - b.day || a.name.localeCompare(b.name, 'ja'));
  v.innerHTML = `
    <div class="card">
      <h2>1. 日程管理シート</h2>
      <p>mainシートを読み取ります。ファイルには書き込みません。</p>
      <label class="file ${state.students.length ? 'set' : ''}" id="lblSch">
        ${state.students.length ? `${state.month}月 ・ 生徒${state.students.length}名 ・ ${state.loadedAt ? new Date(state.loadedAt).toLocaleString('ja-JP', { month: 'numeric', day: 'numeric', hour: '2-digit', minute: '2-digit' }) + ' 読込' : ''}` : 'ファイルを選ぶ（.xlsm / .xlsx）'}
        <input type="file" id="fSch" accept=".xlsm,.xlsx" style="display:none"></label>
    </div>
    <div class="card">
      <h2>2. シフト表</h2>
      <p>出勤人数から受け入れ上限を計算します。ExcelでもPDFでも読めます。</p>
      <label class="file ${state.shiftDays.length ? 'set' : ''}" id="lblSft">
        ${state.shiftDays.length ? `${state.shiftDays.length}日分 ・ ${state.shiftAt ? new Date(state.shiftAt).toLocaleString('ja-JP', { month: 'numeric', day: 'numeric', hour: '2-digit', minute: '2-digit' }) + ' 読込' : ''}` : 'ファイルを選ぶ（.xlsx / .pdf）'}
        <input type="file" id="fSft" accept=".xlsx,.xlsm,.pdf" style="display:none"></label>
    </div>
    <div class="card">
      <h2>3. Excelへ反映する （${lines.length}件）</h2>
      <p>アプリで決めた分です。反映したら「反映済みにする」を押してください。</p>
      ${lines.length ? `<pre class="out" id="outText">${esc(lines.map(p => state.month + '/' + p.day + '\t' + (p.no || p.name)).join('\n'))}</pre>
        <div class="row2" style="margin-bottom:10px">
          <button class="btn sub" id="btnCopy">コピー</button>
          <button class="btn sub" id="btnCsv">CSVで保存</button>
        </div>
        <div id="pendList">${lines.map((p, i) => `<div class="pend"><span class="d">${state.month}/${p.day}</span>
          <span class="n">${esc(p.name)}</span><button data-k="${esc(p.name)}|${p.day}">取消</button></div>`).join('')}</div>
        <button class="btn warnbtn" id="btnClear" style="margin-top:12px">反映済みにする（一覧を空にする）</button>`
      : '<p style="margin:0">まだありません。</p>'}
    </div>
    <div class="card">
      <h2>同期（iPhoneと校舎PC）</h2>
      <p>${syncReady()
        ? (state.syncErr ? '<span style="color:#B23A2E">つながりません：' + esc(state.syncErr) + '</span>'
          : '接続中' + (state.syncAt ? '　最終同期 ' + new Date(state.syncAt).toLocaleString('ja-JP', { month: 'numeric', day: 'numeric', hour: '2-digit', minute: '2-digit' }) : ''))
        : '設定するとiPhoneと校舎PCで同じ内容になります。'}</p>
      <label class="f">ウェブアプリのURL<input type="text" id="inUrl" value="${esc(state.sync.url)}" placeholder="https://script.google.com/macros/s/.../exec"></label>
      <label class="f">合言葉<input type="text" id="inTok" value="${esc(state.sync.token)}" placeholder="Apps Scriptに書いたもの"></label>
      <div class="row2" style="margin-top:8px">
        <button class="btn sub" id="btnSyncTest">接続を確認</button>
        <button class="btn sub" id="btnSyncNow">今すぐ同期</button>
      </div>
    </div>
    <div class="card">
      <h2>設定</h2>
      <div class="row2">
        <label class="f">1人あたりの上限（件）<input type="number" id="inCap" value="${state.cap}" min="1" max="20"></label>
        <label class="f">1人あたりの回数<input type="number" id="inTarget" value="${state.target}" min="1" max="10"></label>
      </div>
      ${Object.keys(state.teams).map(t => `<label class="f">${t}チーム
        <input type="text" data-team="${t}" value="${esc(state.teams[t].join('　'))}"></label>`).join('')}
      <p style="margin:6px 0 0">スタッフ名はシフト表の書き方に合わせてください。全角スペースかカンマ区切りです。</p>
    </div>`;

  $('fSch').addEventListener('change', e => e.target.files[0] && importSchedule(e.target.files[0]));
  $('fSft').addEventListener('change', e => e.target.files[0] && importShift(e.target.files[0]));
  if (lines.length) {
    $('btnCopy').addEventListener('click', () => {
      const t = $('outText').textContent;
      navigator.clipboard ? navigator.clipboard.writeText(t).then(() => toast('コピーしました')) : toast('長押しで選択してコピーしてください');
    });
    $('btnCsv').addEventListener('click', () => {
      const csv = '\uFEFF' + '月,日,マナビス生番号,氏名,チーム\n' + lines.map(p => `${state.month},${p.day},${p.no || ''},"${p.name}",${p.team}`).join('\n');
      const a = document.createElement('a');
      a.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv' }));
      a.download = `面談設定_${state.month}月.csv`;
      document.body.appendChild(a); a.click(); a.remove();
    });
    $('btnClear').addEventListener('click', async () => {
      if (!confirm('Excelへの反映は終わりましたか？一覧を空にします。')) return;
      if (syncReady()) {
        try { await syncCall('done', { ids: state.pending.filter(p => p.id).map(p => p.id) }); }
        catch (e) { if (!confirm('同期先を更新できませんでした（' + e.message + '）。手元の一覧だけ空にしますか？')) return; }
      }
      state.pending = []; await save(); render(); toast('空にしました');
    });
    v.querySelectorAll('.pend button').forEach(b => b.addEventListener('click', async () => {
      const [nm, d] = b.dataset.k.split('|');
      const gone = state.pending.filter(p => p.name === nm && String(p.day) === d);
      state.pending = state.pending.filter(p => !(p.name === nm && String(p.day) === d));
      await save(); render();
      if (syncReady()) {
        const ids = gone.map(p => p.id).filter(Boolean);
        if (ids.length) { try { await syncCall('done', { ids: ids }); } catch (e) { toast('同期先から消せませんでした'); } }
      }
    }));
  }
  $('inUrl').addEventListener('change', async e => { state.sync.url = e.target.value.trim(); await save(); renderHeader(); });
  $('inTok').addEventListener('change', async e => { state.sync.token = e.target.value.trim(); await save(); renderHeader(); });
  $('btnSyncTest').addEventListener('click', async () => {
    state.sync.url = $('inUrl').value.trim(); state.sync.token = $('inTok').value.trim(); await save();
    if (!syncReady()) { toast('URLと合言葉の両方を入れてください'); return; }
    try { const r = await syncCall('ping', {}); state.syncErr = ''; toast('つながりました（未反映 ' + r.count + '件）'); }
    catch (e) { state.syncErr = e.message; toast('つながりません：' + e.message); }
    render();
  });
  $('btnSyncNow').addEventListener('click', () => syncPush(false).then(render));
  $('inCap').addEventListener('change', async e => { state.cap = Math.max(1, Number(e.target.value) || 6); await save(); render(); });
  $('inTarget').addEventListener('change', async e => { state.target = Math.max(1, Number(e.target.value) || 3); await save(); render(); });
  v.querySelectorAll('[data-team]').forEach(inp => inp.addEventListener('change', async e => {
    state.teams[e.target.dataset.team] = e.target.value.split(/[\s,、　]+/).filter(Boolean);
    await save(); render();
  }));
}

async function importSchedule(file) {
  try {
    toast('読み込み中…');
    const wb = XLSX.read(new Uint8Array(await file.arrayBuffer()), { cellDates: true, type: 'array' });
    const r = parseSchedule(wb);
    state.month = r.month; state.days = r.days; state.students = r.students;
    state.noByName = readNoByName(wb);
    state.loadedAt = Date.now();
    if (!state.year) state.year = new Date().getFullYear();
    // Excelに入っている分は、未反映一覧から消す
    const before = state.pending.length;
    state.pending = state.pending.filter(p => {
      const s = r.students.find(x => x.name === p.name);
      return !(s && (s.marks[p.day] === '予定日' || s.marks[p.day] === '済'));
    });
    await save(); render();
    toast(`${r.month}月・${r.students.length}名を読み込みました` + (before - state.pending.length ? `（反映済み${before - state.pending.length}件を消しました）` : ''));
  } catch (e) { alert(e.message); }
}
async function importShift(file) {
  try {
    toast('読み込み中…');
    state.shiftDays = await parseShiftFile(file);
    state.shiftAt = Date.now();
    const y = state.shiftDays.length ? state.shiftDays[0].y : null;
    if (y) state.year = y;
    await save(); render();
    toast(`シフト表 ${state.shiftDays.length}日分を読み込みました`);
  } catch (e) { alert(e.message); }
}

/* ---------- 起動 ---------- */
(async function init() {
  await load();
  if (!state.sync) state.sync = { url: '', token: '' };
  if (!state.noByName) state.noByName = {};
  render();
  autoSync();                                   // 起動時に取り込む
  // 画面に戻ってきたときにも取り込む（別の端末で入れた分を拾う）
  document.addEventListener('visibilitychange', () => { if (!document.hidden) autoSync(); });
})();

})();
</script>
</body>
</html>
