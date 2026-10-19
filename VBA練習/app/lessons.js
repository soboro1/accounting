/* ================================================================
   サンプルデータ ＋ レッスン定義
   ================================================================ */
(function (global) {
'use strict';

/* ---------- 決まった並びのサンプルデータ ---------- */
var seed = 20250912;
function rnd(){ seed = (seed * 1103515245 + 12345) % 2147483648; return seed / 2147483648; }
function pick(a){ return a[Math.floor(rnd() * a.length)]; }

var PRODUCTS = [
  ['ボールペン 黒', '文具', 780],
  ['A4コピー用紙', '用紙', 480],
  ['付箋 75×75', '文具', 320],
  ['トナー 黒', 'トナー', 8900],
  ['インクカートリッジ', 'トナー', 3200],
  ['クリアファイル', 'ファイル', 1280],
  ['リングファイル', 'ファイル', 380],
  ['電卓 12桁', '備品', 1980]
];
var BRANCHES = ['東京', '大阪', '名古屋'];
var STAFF = ['佐藤 健一', '鈴木 美咲', '高橋 大輔', '田中 由美', '伊藤 翔'];

function buildSales(){
  seed = 20250912;
  var rows = [];
  for (var i = 0; i < 40; i++){
    var p = pick(PRODUCTS);
    var qty = [1,2,3,5,10,20][Math.floor(rnd()*6)];
    rows.push([
      new Date(2025, Math.floor(i/14), (i % 28) + 1),
      pick(BRANCHES), pick(STAFF), p[0], p[1], p[2], qty, p[2] * qty
    ]);
  }
  return rows;
}
var INVENTORY = [
  ['ボールペン 黒', 120, 30], ['A4コピー用紙', 18, 30], ['付箋 75×75', 0, 20],
  ['トナー 黒', 45, 20], ['インクカートリッジ', 30, 30], ['クリアファイル', 210, 50],
  ['リングファイル', 7, 20], ['電卓 12桁', 0, 10], ['デスクマット', 64, 20], ['卓上ライト', 25, 25]
];
var STAFFLIST = [
  ['E001','佐藤 健一','営業1課', new Date(2012,3,1), 520000],
  ['E002','鈴木 美咲','営業1課', new Date(2016,9,1), 385000],
  ['E003','高橋 大輔','営業2課', new Date(2009,3,1), 610000],
  ['E004','田中 由美','管理部',  new Date(2021,3,1), 298000],
  ['E005','伊藤 翔',  '営業2課', new Date(2019,7,16), 342000],
  ['E006','渡辺 彩',  '物流部',  new Date(2023,3,1), 265000],
  ['E007','山本 直樹','営業1課', new Date(2014,0,6), 455000],
  ['E008','中村 恵子','管理部',  new Date(2011,3,1), 498000],
  ['E009','小林 拓也','物流部',  new Date(2022,9,1), 288000],
  ['E010','加藤 真希','営業2課', new Date(2018,3,2), 370000],
  ['E011','吉田 亮',  '営業1課', new Date(2020,5,1), 412000],
  ['E012','松本 沙織','管理部',  new Date(2017,3,3), 336000]
];
var DIRTY = [
  ['C001', ' 株式会社アルファ  ', '１２０，０００'],
  ['C002', 'ベータ商事(株)',      '85000'],
  ['C003', '株式会社ガンマ',      '240,000'],
  ['C004', 'デルタ工業  株式会社','56,000'],
  ['C005', 'ｲﾌﾟｼﾛﾝ物産',          '３３０，０００'],
  ['C006', 'ゼータ建設株式会社',  '  1450000  '],
  ['C007', 'イータサービス',      '７８０００'],
  ['C008', 'シータ物流㈱',        '92,000'],
  ['C009', 'ヨタ電機株式会社',    '  305，000 ']
];

function buildWorkbook(){
  var RT = global.VBARuntime;
  function sheetFrom(name, headers, rows, fmts){
    var s = new RT.Sheet(name);
    headers.forEach(function(h, i){
      s.setV(1, i+1, h);
      var c = s.cell(1, i+1, true); c.bold = true; c.fill = 0xE8EFE9; c.header = true;
    });
    rows.forEach(function(r, ri){
      r.forEach(function(v, ci){
        s.setV(ri+2, ci+1, v);
        if (fmts && fmts[ci]) s.cell(ri+2, ci+1, true).fmt = fmts[ci];
      });
    });
    return s;
  }
  var sales = sheetFrom('売上明細',
    ['日付','支店','担当者','商品名','カテゴリ','単価','数量','金額'],
    buildSales(), { 0:'yyyy/mm/dd', 5:'#,##0', 7:'#,##0' });
  var master = sheetFrom('商品マスタ', ['商品名','カテゴリ','単価'],
    PRODUCTS.map(function(p){ return p.slice(); }), { 2:'#,##0' });
  var inv = sheetFrom('在庫', ['商品名','在庫数','発注点','判定'],
    INVENTORY.map(function(r){ return r.concat(['']); }), { 1:'#,##0' });
  var staff = sheetFrom('社員名簿', ['社員番号','氏名','部署','入社日','基本給','ランク'],
    STAFFLIST.map(function(r){ return r.concat(['']); }), { 3:'yyyy/mm/dd', 4:'#,##0' });
  var dirty = sheetFrom('要整形データ', ['顧客番号','会社名','契約金額(文字列)','数値'],
    DIRTY.map(function(r){ return r.concat(['']); }), null);
  var out = new RT.Sheet('集計結果');
  return new RT.Workbook([sales, master, inv, staff, dirty, out]);
}

/* ---------- 採点ヘルパー ---------- */
var EPS = 0.001;
function num(v){ return typeof v === 'number' ? v : (v === '' || v === undefined || v === null || (v && v.__empty)) ? NaN : Number(String(v).replace(/,/g,'')); }
function eqNum(a, b){ var x = num(a); return !isNaN(x) && Math.abs(x - b) < EPS; }
function txt(v){
  if (v === undefined || v === null || (v && v.__empty)) return '';
  if (v instanceof Date) return global.VBARuntime.toStr(v);
  return String(v).trim();
}
function cell(ctx, sheetName, r, c){ var s = ctx.wb.byName(sheetName); return s ? s.getV(r, c) : undefined; }
function fail(m){ return { ok: false, msg: m }; }
function pass(m){ return { ok: true, msg: m || '正解です。' }; }
function msgboxText(ctx){ return ctx.out.filter(function(o){ return o.kind === 'msgbox'; }).map(function(o){ return o.text; }).join('\n'); }
function baseSales(ctx){
  var s = ctx.base.byName('売上明細'), rows = [];
  for (var r = 2; r <= 41; r++) rows.push({ date: s.getV(r,1), branch: s.getV(r,2), staff: s.getV(r,3), item: s.getV(r,4), cat: s.getV(r,5), price: s.getV(r,6), qty: s.getV(r,7), amt: s.getV(r,8) });
  return rows;
}

/* ---------- レッスン ---------- */
var LESSONS = [
{
  id:'l01', level:1, title:'はじめてのマクロ', sheet:'売上明細',
  goal:'<code>MsgBox</code> で <b>Hello VBA</b> と表示してください。',
  body:'<p>VBA の処理は必ず <code>Sub 名前()</code> と <code>End Sub</code> の間に書きます。この2行の「入れもの」の中に命令を並べていくのが、すべてのマクロの基本の形です。</p><p><code>MsgBox</code> はメッセージを表示する命令。文字列は必ず <b>半角のダブルクォーテーション</b> で囲みます。</p>',
  hints:['<code>MsgBox "表示したい文字"</code> と書きます。','日本語入力のまま <code>”</code>（全角）を打つとエラーになります。半角の <code>"</code> で囲んでください。'],
  starter:'Sub 練習01()\n    \' ここに MsgBox を書く\n\nEnd Sub',
  solution:'Sub 練習01()\n    MsgBox "Hello VBA"\nEnd Sub',
  check:function(ctx){
    var t = msgboxText(ctx);
    if (!t) return fail('MsgBox が1度も表示されませんでした。');
    if (!/hello\s*vba/i.test(t)) return fail('表示されたのは「' + t + '」でした。「Hello VBA」と表示してください。');
    return pass('メッセージが表示されました。これがマクロの最小単位です。');
  }
},
{
  id:'l02', level:1, title:'セルに書き込む', sheet:'集計結果',
  goal:'「集計結果」シートに、<code>A1</code>＝<b>商品名</b>、<code>B1</code>＝<b>金額</b>、<code>A2</code>＝<b>ボールペン 黒</b>、<code>B2</code>＝<b>780</b> を書き込んでください。',
  body:'<p>セルを指すには <code>Range("A1")</code>。書き込むのは <code>.Value</code> プロパティです。</p><p>どのシートのセルかを必ず指定するのが実務の鉄則。<code>ws.Range("A1")</code> のように、シートを変数に入れてから使います。シート指定を省くと「今開いているシート」が対象になり、事故のもとです。</p>',
  hints:['<code>Dim ws As Worksheet</code> → <code>Set ws = ThisWorkbook.Worksheets("集計結果")</code>','オブジェクトを変数に入れるときは <code>=</code> ではなく <code>Set</code> を使います。','文字列は <code>"</code> で囲み、数値は囲みません。'],
  starter:'Sub 練習02()\n    Dim ws As Worksheet\n    Set ws = ThisWorkbook.Worksheets("集計結果")\n\n    ws.Range("A1").Value = "商品名"\n    \' 残り3つを書く\n\nEnd Sub',
  solution:'Sub 練習02()\n    Dim ws As Worksheet\n    Set ws = ThisWorkbook.Worksheets("集計結果")\n\n    ws.Range("A1").Value = "商品名"\n    ws.Range("B1").Value = "金額"\n    ws.Range("A2").Value = "ボールペン 黒"\n    ws.Range("B2").Value = 780\nEnd Sub',
  check:function(ctx){
    var want = [['A1',1,1,'商品名'],['B1',1,2,'金額'],['A2',2,1,'ボールペン 黒']];
    for (var i=0;i<want.length;i++){
      var got = txt(cell(ctx,'集計結果',want[i][1],want[i][2]));
      if (got !== want[i][3]) return fail(want[i][0] + ' が「' + (got||'空') + '」です。「' + want[i][3] + '」にしてください。');
    }
    var b2 = cell(ctx,'集計結果',2,2);
    if (!eqNum(b2, 780)) return fail('B2 が「' + txt(b2) + '」です。数値の 780 を入れてください（"780" と書くと文字列になります）。');
    if (typeof b2 !== 'number') return fail('B2 が文字列になっています。ダブルクォートを外して 780 と書いてください。');
    return pass('4つのセルに書き込めました。');
  }
},
{
  id:'l03', level:1, title:'セルを読んで計算する', sheet:'売上明細',
  goal:'「売上明細」の2行目の<b>単価</b>（F2）と<b>数量</b>（G2）を読み、金額を「集計結果」の <code>A1</code> に、税込（10%・小数切り捨て）を <code>A2</code> に書いてください。',
  body:'<p>セルは書くだけでなく読めます。<code>ws.Range("F2").Value</code> と書けばその値が取り出せます。</p><p>切り捨ては <code>Int()</code>。<code>Int(1234 * 1.1)</code> のように使います。</p>',
  hints:['単価 × 数量 が金額です。','<code>Int(金額 * 1.1)</code> で税込。','読むシートと書くシートが違うので、変数を2つ用意すると分かりやすくなります。'],
  starter:'Sub 練習03()\n    Dim wsS As Worksheet, wsO As Worksheet\n    Set wsS = ThisWorkbook.Worksheets("売上明細")\n    Set wsO = ThisWorkbook.Worksheets("集計結果")\n\n\nEnd Sub',
  solution:'Sub 練習03()\n    Dim wsS As Worksheet, wsO As Worksheet\n    Dim price As Long, qty As Long, amount As Long\n    Set wsS = ThisWorkbook.Worksheets("売上明細")\n    Set wsO = ThisWorkbook.Worksheets("集計結果")\n\n    price = wsS.Range("F2").Value\n    qty = wsS.Range("G2").Value\n    amount = price * qty\n\n    wsO.Range("A1").Value = amount\n    wsO.Range("A2").Value = Int(amount * 1.1)\nEnd Sub',
  check:function(ctx){
    var s = ctx.base.byName('売上明細');
    var amt = s.getV(2,6) * s.getV(2,7);
    var a1 = cell(ctx,'集計結果',1,1), a2 = cell(ctx,'集計結果',2,1);
    if (!eqNum(a1, amt)) return fail('A1 が「' + txt(a1) + '」です。単価 ' + s.getV(2,6) + ' × 数量 ' + s.getV(2,7) + ' = ' + amt + ' になるはずです。');
    if (!eqNum(a2, Math.floor(amt*1.1))) return fail('A2 が「' + txt(a2) + '」です。Int(' + amt + ' * 1.1) = ' + Math.floor(amt*1.1) + ' になるはずです。');
    return pass('セルを読んで計算し、別シートに書けました。');
  }
},
{
  id:'l04', level:1, title:'変数と型', sheet:'集計結果',
  goal:'変数を使って「売上明細」2行目の金額を求め、<code>MsgBox</code> で <b>単価 × 数量 = 金額</b> の形（例: <code>780 x 10 = 7800</code>）で表示してください。',
  body:'<p><code>Dim price As Long</code> のように、変数は<b>使う前に宣言</b>します。<code>Option Explicit</code> を書いておくと宣言し忘れをエラーで教えてくれるので、必ず付けてください。タイプミスを実行前に潰せます。</p><p>型は、整数なら <code>Long</code>、文字なら <code>String</code>、小数が出るなら <code>Double</code>。金額の合計は桁が大きくなるので <code>Currency</code> か <code>Double</code> が安全です。</p><p>文字列の連結は <code>&amp;</code> を使います。<code>price &amp; " x " &amp; qty</code> のように。</p>',
  hints:['<code>MsgBox price &amp; " x " &amp; qty &amp; " = " &amp; amount</code>','<code>&amp;</code> の前後には半角スペースが必要です（<code>price&amp;qty</code> はエラー）。'],
  starter:'Option Explicit\n\nSub 練習04()\n    Dim ws As Worksheet\n    Dim price As Long, qty As Long, amount As Long\n    Set ws = ThisWorkbook.Worksheets("売上明細")\n\n\nEnd Sub',
  solution:'Option Explicit\n\nSub 練習04()\n    Dim ws As Worksheet\n    Dim price As Long, qty As Long, amount As Long\n    Set ws = ThisWorkbook.Worksheets("売上明細")\n\n    price = ws.Range("F2").Value\n    qty = ws.Range("G2").Value\n    amount = price * qty\n\n    MsgBox price & " x " & qty & " = " & amount\nEnd Sub',
  check:function(ctx){
    var s = ctx.base.byName('売上明細');
    var p = s.getV(2,6), q = s.getV(2,7), a = p*q;
    var t = msgboxText(ctx);
    if (!t) return fail('MsgBox が表示されませんでした。');
    var nums = (t.match(/\d+/g) || []).map(Number);
    if (nums.indexOf(p) < 0 || nums.indexOf(q) < 0 || nums.indexOf(a) < 0)
      return fail('表示は「' + t + '」でした。' + p + '、' + q + '、' + a + ' の3つが含まれるようにしてください。');
    return pass('変数に入れて計算し、連結して表示できました。');
  }
},
{
  id:'l05', level:2, title:'If で条件分岐', sheet:'在庫',
  goal:'「在庫」シートの <b>D列（判定）</b> を2行目から11行目まで埋めてください。<br>在庫数 = 0 → <b>欠品</b>／在庫数 ≦ 発注点 → <b>要発注</b>／それ以外 → <b>OK</b>',
  body:'<p>条件分岐は <code>If 〜 Then</code>／<code>ElseIf 〜 Then</code>／<code>Else</code>／<code>End If</code>。</p><p><b>順番が結果を変えます。</b>「0 かどうか」を先に見ないと、在庫0の行が「要発注」に吸い込まれてしまいます（0 ≦ 発注点 なので）。条件は<b>狭いものから先に</b>書くのが鉄則です。</p><p>10行あるので <code>For i = 2 To 11</code> で回します。<code>Cells(i, 2)</code> は「i 行目の2列目」＝B列です。</p>',
  hints:['<code>For i = 2 To 11</code> ... <code>Next i</code> でループ。','<code>ws.Cells(i, 2).Value</code> が在庫数、<code>ws.Cells(i, 3).Value</code> が発注点。','書き込むのは <code>ws.Cells(i, 4).Value</code>（D列）。'],
  starter:'Option Explicit\n\nSub 練習05()\n    Dim ws As Worksheet\n    Dim i As Long\n    Set ws = ThisWorkbook.Worksheets("在庫")\n\n    For i = 2 To 11\n\n    Next i\nEnd Sub',
  solution:'Option Explicit\n\nSub 練習05()\n    Dim ws As Worksheet\n    Dim i As Long\n    Set ws = ThisWorkbook.Worksheets("在庫")\n\n    For i = 2 To 11\n        If ws.Cells(i, 2).Value = 0 Then\n            ws.Cells(i, 4).Value = "欠品"\n        ElseIf ws.Cells(i, 2).Value <= ws.Cells(i, 3).Value Then\n            ws.Cells(i, 4).Value = "要発注"\n        Else\n            ws.Cells(i, 4).Value = "OK"\n        End If\n    Next i\nEnd Sub',
  check:function(ctx){
    var b = ctx.base.byName('在庫');
    for (var r=2;r<=11;r++){
      var stock = b.getV(r,2), reorder = b.getV(r,3);
      var want = stock === 0 ? '欠品' : (stock <= reorder ? '要発注' : 'OK');
      var got = txt(cell(ctx,'在庫',r,4));
      if (got !== want) return fail(r + '行目（' + b.getV(r,1) + '：在庫' + stock + '／発注点' + reorder + '）が「' + (got||'空') + '」です。正しくは「' + want + '」。' + (stock===0&&got==='要発注' ? ' 在庫0の判定を先に書いてください。' : ''));
    }
    return pass('10行すべて正しく判定できました。条件の順番が効いています。');
  }
},
{
  id:'l06', level:2, title:'For で合計する', sheet:'売上明細',
  goal:'「売上明細」の <b>H列（金額）</b> を2行目から41行目まで合計し、「集計結果」の <code>B1</code> に書いてください。',
  body:'<p>合計の型は <code>For i = 2 To 41</code> でループし、変数に足し込むだけです。</p><pre class="mini">total = total + ws.Cells(i, 8).Value</pre><p>この <code>total = total + 〜</code> は数学の等式ではなく「<b>右辺を計算して左辺に入れ直す</b>」という命令です。ここが腑に落ちるとループが書けるようになります。</p><p>合計を入れる変数は <code>Long</code> だと桁あふれの危険があるので <code>Currency</code> か <code>Double</code> にします。</p>',
  hints:['<code>Dim total As Currency</code>','H列は8列目なので <code>ws.Cells(i, 8)</code>。','ループの外で <code>wsO.Range("B1").Value = total</code>。'],
  starter:'Option Explicit\n\nSub 練習06()\n    Dim wsS As Worksheet, wsO As Worksheet\n    Dim i As Long, total As Currency\n    Set wsS = ThisWorkbook.Worksheets("売上明細")\n    Set wsO = ThisWorkbook.Worksheets("集計結果")\n\n\nEnd Sub',
  solution:'Option Explicit\n\nSub 練習06()\n    Dim wsS As Worksheet, wsO As Worksheet\n    Dim i As Long, total As Currency\n    Set wsS = ThisWorkbook.Worksheets("売上明細")\n    Set wsO = ThisWorkbook.Worksheets("集計結果")\n\n    For i = 2 To 41\n        total = total + wsS.Cells(i, 8).Value\n    Next i\n\n    wsO.Range("B1").Value = total\nEnd Sub',
  check:function(ctx){
    var want = baseSales(ctx).reduce(function(a,r){ return a + r.amt; }, 0);
    var got = cell(ctx,'集計結果',1,2);
    if (!eqNum(got, want)) return fail('B1 が「' + txt(got) + '」です。正解は ' + want.toLocaleString() + ' です。' + (isNaN(num(got)) ? ' ループの外で書き込んでいますか？' : ''));
    return pass('40行を足し上げられました。合計は ' + want.toLocaleString() + ' 円。');
  }
},
{
  id:'l07', level:2, title:'Do While で数える', sheet:'売上明細',
  goal:'<code>Do While</code> を使って、「売上明細」で <b>数量（G列）が5以上</b> の行数を数え、「集計結果」の <code>B2</code> に書いてください。',
  body:'<p><code>For</code> は回数が決まっているとき、<code>Do While</code> は<b>いつ終わるか分からないとき</b>に使います。ここでは「A列が空になるまで」進めます。</p><pre class="mini">i = 2\nDo While ws.Cells(i, 1).Value &lt;&gt; ""\n    \' 処理\n    i = i + 1\nLoop</pre><p><b>最後の <code>i = i + 1</code> を忘れると無限ループです。</b>実務では <code>Ctrl + Break</code> で止めます（このアプリは自動で止めます）。</p>',
  hints:['数えるための変数 <code>Dim cnt As Long</code> を用意します。','<code>If ws.Cells(i, 7).Value >= 5 Then cnt = cnt + 1</code>','<code>&lt;&gt;</code> は「等しくない」。<code>""</code> は空文字です。'],
  starter:'Option Explicit\n\nSub 練習07()\n    Dim wsS As Worksheet, wsO As Worksheet\n    Dim i As Long, cnt As Long\n    Set wsS = ThisWorkbook.Worksheets("売上明細")\n    Set wsO = ThisWorkbook.Worksheets("集計結果")\n\n    i = 2\n    Do While wsS.Cells(i, 1).Value <> ""\n\n    Loop\n\nEnd Sub',
  solution:'Option Explicit\n\nSub 練習07()\n    Dim wsS As Worksheet, wsO As Worksheet\n    Dim i As Long, cnt As Long\n    Set wsS = ThisWorkbook.Worksheets("売上明細")\n    Set wsO = ThisWorkbook.Worksheets("集計結果")\n\n    i = 2\n    Do While wsS.Cells(i, 1).Value <> ""\n        If wsS.Cells(i, 7).Value >= 5 Then cnt = cnt + 1\n        i = i + 1\n    Loop\n\n    wsO.Range("B2").Value = cnt\nEnd Sub',
  check:function(ctx){
    var want = baseSales(ctx).filter(function(r){ return r.qty >= 5; }).length;
    var got = cell(ctx,'集計結果',2,2);
    if (!eqNum(got, want)) return fail('B2 が「' + txt(got) + '」です。正解は ' + want + ' 件です。');
    return pass('Do While で ' + want + ' 件を数えられました。');
  }
},
{
  id:'l08', level:2, title:'Select Case', sheet:'社員名簿',
  goal:'「社員名簿」の <b>F列（ランク）</b> を、基本給（E列）から埋めてください。<br>50万以上＝<b>A</b>／40万台＝<b>B</b>／30万台＝<b>C</b>／それ未満＝<b>D</b>',
  body:'<p>同じ変数を何度も比べるときは、<code>If / ElseIf</code> を重ねるより <code>Select Case</code> のほうが読みやすくなります。</p><pre class="mini">Select Case salary\n    Case Is &gt;= 500000\n        rank = "A"\n    Case 400000 To 499999\n        rank = "B"\n    Case Else\n        rank = "D"\nEnd Select</pre><p><code>Case Is &gt;= 500000</code> の <code>Is</code> は「比べる対象＝Select Case に書いた値」を指すキーワードです。</p>',
  hints:['12人なので <code>For i = 2 To 13</code>。','基本給は E列＝5列目、ランクは F列＝6列目。','<code>Case 300000 To 399999</code> のように範囲も書けます。'],
  starter:'Option Explicit\n\nSub 練習08()\n    Dim ws As Worksheet\n    Dim i As Long, salary As Long\n    Set ws = ThisWorkbook.Worksheets("社員名簿")\n\n    For i = 2 To 13\n        salary = ws.Cells(i, 5).Value\n\n    Next i\nEnd Sub',
  solution:'Option Explicit\n\nSub 練習08()\n    Dim ws As Worksheet\n    Dim i As Long, salary As Long\n    Set ws = ThisWorkbook.Worksheets("社員名簿")\n\n    For i = 2 To 13\n        salary = ws.Cells(i, 5).Value\n        Select Case salary\n            Case Is >= 500000\n                ws.Cells(i, 6).Value = "A"\n            Case 400000 To 499999\n                ws.Cells(i, 6).Value = "B"\n            Case 300000 To 399999\n                ws.Cells(i, 6).Value = "C"\n            Case Else\n                ws.Cells(i, 6).Value = "D"\n        End Select\n    Next i\nEnd Sub',
  check:function(ctx){
    var b = ctx.base.byName('社員名簿');
    for (var r=2;r<=13;r++){
      var s = b.getV(r,5);
      var want = s >= 500000 ? 'A' : s >= 400000 ? 'B' : s >= 300000 ? 'C' : 'D';
      var got = txt(cell(ctx,'社員名簿',r,6)).toUpperCase();
      if (got !== want) return fail(r + '行目（' + b.getV(r,2) + '：' + s.toLocaleString() + '円）が「' + (got||'空') + '」です。正しくは「' + want + '」。');
    }
    return pass('12人すべて正しくランク付けできました。');
  }
},
{
  id:'l09', level:3, title:'最終行を求める', sheet:'売上明細',
  goal:'<code>End(xlUp)</code> で「売上明細」の最終行番号を求め、「集計結果」の <code>B1</code> に<b>最終行番号</b>、<code>B2</code> に<b>データ件数</b>（見出しを除く）を書いてください。',
  body:'<p>実務のマクロで <code>For i = 2 To 41</code> と行数を直書きするのは事故のもとです。データが増えた翌月に壊れます。<b>行数は毎回コードで求めます。</b></p><pre class="mini">lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row</pre><p>これは「A列のいちばん下のセルから <kbd>Ctrl</kbd>+<kbd>↑</kbd> を押す」という操作そのものです。<code>ws.Rows.Count</code> はシートの総行数（1,048,576）。そこから上に詰めた着地点の行番号が <code>.Row</code> です。</p><p>見出しが1行あるので、データ件数は <code>lastRow - 1</code> になります。',
  hints:['<code>Dim lastRow As Long</code>','列は文字でも番号でも指定できます（<code>"A"</code> でも <code>1</code> でも可）。','<code>.End(xlUp)</code> が返すのは<b>セル</b>なので、行番号を取るには <code>.Row</code> を付けます。'],
  starter:'Option Explicit\n\nSub 練習09()\n    Dim wsS As Worksheet, wsO As Worksheet\n    Dim lastRow As Long\n    Set wsS = ThisWorkbook.Worksheets("売上明細")\n    Set wsO = ThisWorkbook.Worksheets("集計結果")\n\n\nEnd Sub',
  solution:'Option Explicit\n\nSub 練習09()\n    Dim wsS As Worksheet, wsO As Worksheet\n    Dim lastRow As Long\n    Set wsS = ThisWorkbook.Worksheets("売上明細")\n    Set wsO = ThisWorkbook.Worksheets("集計結果")\n\n    lastRow = wsS.Cells(wsS.Rows.Count, "A").End(xlUp).Row\n\n    wsO.Range("B1").Value = lastRow\n    wsO.Range("B2").Value = lastRow - 1\nEnd Sub',
  check:function(ctx){
    var b1 = cell(ctx,'集計結果',1,2), b2 = cell(ctx,'集計結果',2,2);
    if (!eqNum(b1, 41)) return fail('B1 が「' + txt(b1) + '」です。最終行は 41 です。' + (eqNum(b1,1048576) ? ' Rows.Count のセルから .End(xlUp) していますか？' : ''));
    if (!eqNum(b2, 40)) return fail('B2 が「' + txt(b2) + '」です。見出しを除くと 40 件です。');
    return pass('行数を直書きせずに求められました。これが実務の基本形です。');
  }
},
{
  id:'l10', level:3, title:'書式を変える', sheet:'売上明細',
  goal:'「売上明細」の見出し行 <code>A1:H1</code> を<b>太字</b>にし、<b>背景色 RGB(29, 122, 77)</b>・<b>文字色 白</b>にしてください。さらに <b>H列の2〜41行目</b> に表示形式 <code>#,##0</code> を設定してください。',
  body:'<p>見た目を変えるプロパティは、どれも「範囲 → 何を → どうする」の形です。</p><pre class="mini">.Font.Bold = True\n.Font.Color = RGB(255, 255, 255)\n.Interior.Color = RGB(29, 122, 77)\n.NumberFormat = "#,##0"</pre><p>同じ範囲に何度も触るときは <code>With</code> でまとめると、範囲の指定を1回で済ませられます。実行も速くなります。</p><pre class="mini">With ws.Range("A1:H1")\n    .Font.Bold = True\n    .Interior.Color = RGB(29, 122, 77)\nEnd With</pre><p><code>RGB(赤, 緑, 青)</code> は各 0〜255。白は <code>RGB(255,255,255)</code> です。</p>',
  hints:['<code>With</code> の中では先頭の <code>.</code> を省略できません（<code>.Font.Bold</code>）。','表示形式は <code>ws.Range("H2:H41").NumberFormat = "#,##0"</code>。','<code>#,##0</code> は「3桁区切り・小数なし」の意味です。'],
  starter:'Option Explicit\n\nSub 練習10()\n    Dim ws As Worksheet\n    Set ws = ThisWorkbook.Worksheets("売上明細")\n\n    With ws.Range("A1:H1")\n\n    End With\n\nEnd Sub',
  solution:'Option Explicit\n\nSub 練習10()\n    Dim ws As Worksheet\n    Set ws = ThisWorkbook.Worksheets("売上明細")\n\n    With ws.Range("A1:H1")\n        .Font.Bold = True\n        .Font.Color = RGB(255, 255, 255)\n        .Interior.Color = RGB(29, 122, 77)\n    End With\n\n    ws.Range("H2:H41").NumberFormat = "#,##0"\nEnd Sub',
  check:function(ctx){
    var s = ctx.wb.byName('売上明細');
    for (var c=1;c<=8;c++){
      var x = s.cell(1,c,false);
      if (!x || !x.bold) return fail(global.VBARuntime.colLetter(c) + '1 が太字になっていません。');
      if (!x || x.fill !== 29 + 122*256 + 77*65536) return fail(global.VBARuntime.colLetter(c) + '1 の背景色が違います。RGB(29, 122, 77) にしてください。');
      if (!x || x.fcolor !== 255 + 255*256 + 255*65536) return fail(global.VBARuntime.colLetter(c) + '1 の文字色が白になっていません。RGB(255, 255, 255) です。');
    }
    for (var r=2;r<=41;r++){
      var y = s.cell(r,8,false);
      if (!y || y.fmt !== '#,##0') return fail('H' + r + ' の表示形式が「' + ((y&&y.fmt)||'なし') + '」です。"#,##0" にしてください。');
    }
    return pass('見出しと表示形式が整いました。With を使うと範囲指定は1回で済みます。');
  }
},
{
  id:'l11', level:3, title:'マスタから転記する', sheet:'在庫',
  goal:'「在庫」シートの <code>E1</code> に <b>単価</b> と書き、<code>E2:E11</code> に「商品マスタ」から商品名をキーに単価を転記してください。マスタに無い商品は <b>―</b>（全角ダッシュ）にしてください。',
  body:'<p>VLOOKUP をマクロでやる場面です。素直に書くなら二重ループ（探す側 × マスタ側）。</p><p><code>WorksheetFunction.VLookup</code> は見つからないと<b>実行時エラーで止まります</b>。実務では <code>Application.VLookup</code> を <code>Variant</code> で受けて <code>IsError</code> で判定するのが定石です。</p><pre class="mini">Dim v As Variant\nv = Application.VLookup(key, wsM.Range("A2:C9"), 3, False)\nIf IsError(v) Then ... </pre><p>「デスクマット」と「卓上ライト」はマスタにありません。ここで落ちないコードを書くのが今回の要点です。</p>',
  hints:['マスタは A列＝商品名、C列＝単価（3列目）。','<code>False</code> は「完全一致で探す」の意味。省略すると別物が返ります。','二重ループで書いてもかまいません。<code>Exit For</code> で見つけ次第抜けると速くなります。'],
  starter:'Option Explicit\n\nSub 練習11()\n    Dim wsI As Worksheet, wsM As Worksheet\n    Dim i As Long, v As Variant\n    Set wsI = ThisWorkbook.Worksheets("在庫")\n    Set wsM = ThisWorkbook.Worksheets("商品マスタ")\n\n    wsI.Range("E1").Value = "単価"\n\n    For i = 2 To 11\n\n    Next i\nEnd Sub',
  solution:'Option Explicit\n\nSub 練習11()\n    Dim wsI As Worksheet, wsM As Worksheet\n    Dim i As Long, v As Variant\n    Set wsI = ThisWorkbook.Worksheets("在庫")\n    Set wsM = ThisWorkbook.Worksheets("商品マスタ")\n\n    wsI.Range("E1").Value = "単価"\n\n    For i = 2 To 11\n        v = Application.VLookup(wsI.Cells(i, 1).Value, wsM.Range("A2:C9"), 3, False)\n        If IsError(v) Then\n            wsI.Cells(i, 5).Value = "―"\n        Else\n            wsI.Cells(i, 5).Value = v\n        End If\n    Next i\nEnd Sub',
  check:function(ctx){
    if (txt(cell(ctx,'在庫',1,5)) !== '単価') return fail('E1 に「単価」と書いてください。');
    var m = ctx.base.byName('商品マスタ'), map = {};
    for (var r=2;r<=9;r++) map[m.getV(r,1)] = m.getV(r,3);
    var inv = ctx.base.byName('在庫');
    for (var i=2;i<=11;i++){
      var name = inv.getV(i,1), got = cell(ctx,'在庫',i,5);
      if (map[name] === undefined){
        if (txt(got) !== '―') return fail(i + '行目「' + name + '」はマスタに無い商品です。E列を「―」にしてください（今は「' + (txt(got)||'空') + '」）。');
      } else if (!eqNum(got, map[name])){
        return fail(i + '行目「' + name + '」の単価が「' + txt(got) + '」です。正しくは ' + map[name] + ' です。');
      }
    }
    return pass('見つからない行で落ちずに転記できました。');
  }
},
{
  id:'l12', level:3, title:'支店別に集計する', sheet:'集計結果',
  goal:'「売上明細」を支店ごとに集計し、「集計結果」の <code>A列</code> に支店名、<code>B列</code> に売上合計を、<b>1行目から</b>書き出してください（順番は問いません）。支店は 東京・大阪・名古屋 の3つです。',
  body:'<p>「支店の一覧を回す」×「明細を回す」の<b>二重ループ</b>です。外側で支店を1つ決め、内側で明細を全部見て、その支店の行だけ足す。これを支店の数だけ繰り返します。</p><pre class="mini">branches = Array("東京", "大阪", "名古屋")\nFor b = 0 To UBound(branches)\n    total = 0\n    For i = 2 To lastRow\n        If ... Then total = total + ...\n    Next i\nNext b</pre><p><b>内側に入る前に <code>total = 0</code> に戻すのを忘れないでください。</b>戻し忘れは前の支店の合計が積み上がる、集計マクロで最も多いバグです。</p><p>ちなみにこの書き方は 40行 × 3支店 = 120回セルを読んでいます。Lv4 でこれを一気に速くします。</p>',
  hints:['<code>Array("東京", "大阪", "名古屋")</code> は <code>Variant</code> 型の変数で受けます。','<code>UBound(branches)</code> が配列の最後の添字（0始まりなので 2）。','書き出す行は <code>b + 1</code> 行目になります。'],
  starter:'Option Explicit\n\nSub 練習12()\n    Dim wsS As Worksheet, wsO As Worksheet\n    Dim branches As Variant\n    Dim b As Long, i As Long, total As Currency\n    Set wsS = ThisWorkbook.Worksheets("売上明細")\n    Set wsO = ThisWorkbook.Worksheets("集計結果")\n\n    branches = Array("東京", "大阪", "名古屋")\n\n    For b = 0 To UBound(branches)\n\n    Next b\nEnd Sub',
  solution:'Option Explicit\n\nSub 練習12()\n    Dim wsS As Worksheet, wsO As Worksheet\n    Dim branches As Variant\n    Dim b As Long, i As Long, total As Currency\n    Set wsS = ThisWorkbook.Worksheets("売上明細")\n    Set wsO = ThisWorkbook.Worksheets("集計結果")\n\n    branches = Array("東京", "大阪", "名古屋")\n\n    For b = 0 To UBound(branches)\n        total = 0\n        For i = 2 To 41\n            If wsS.Cells(i, 2).Value = branches(b) Then\n                total = total + wsS.Cells(i, 8).Value\n            End If\n        Next i\n        wsO.Cells(b + 1, 1).Value = branches(b)\n        wsO.Cells(b + 1, 2).Value = total\n    Next b\nEnd Sub',
  check:function(ctx){
    var want = {}; BRANCHES.forEach(function(b){ want[b] = 0; });
    baseSales(ctx).forEach(function(r){ want[r.branch] += r.amt; });
    var s = ctx.wb.byName('集計結果'), seen = {};
    for (var r=1;r<=10;r++){
      var name = txt(s.getV(r,1));
      if (!name) continue;
      if (want[name] === undefined) return fail(r + '行目に「' + name + '」があります。支店は 東京・大阪・名古屋 の3つです。');
      if (seen[name]) return fail('「' + name + '」が2回書き出されています。');
      seen[name] = true;
      if (!eqNum(s.getV(r,2), want[name])) return fail('「' + name + '」の合計が「' + txt(s.getV(r,2)) + '」です。正しくは ' + want[name].toLocaleString() + ' です。合計をループの頭で 0 に戻していますか？');
    }
    var missing = BRANCHES.filter(function(b){ return !seen[b]; });
    if (missing.length) return fail('「' + missing.join('」「') + '」が書き出されていません。');
    return pass('3支店とも一致しました。合計のリセット位置が正しく書けています。');
  }
},
{
  id:'l13', level:4, title:'配列で高速化する', sheet:'売上明細',
  goal:'「売上明細」の <code>I1</code> に <b>税込金額</b> と書き、<code>I2:I41</code> に H列の税込（<code>Int(金額 * 1.1)</code>）を入れてください。<b>セルを1つずつ触らず、配列で一括読み込み・一括書き戻し</b>で行ってください。',
  body:'<p>VBA が遅い最大の原因は「セルを1つずつ触ること」です。セルアクセスは1回ごとに Excel 本体とやり取りが発生します。</p><p>速くする型は決まっています。<b>Range → 配列 → 計算 → Range に戻す</b>。読むのも書くのも1回だけ。</p><pre class="mini">Dim src As Variant, dst() As Variant\nsrc = ws.Range("H2:H41").Value   \' ← 読むのは1回\nReDim dst(1 To UBound(src, 1), 1 To 1)\nFor i = 1 To UBound(src, 1)\n    dst(i, 1) = Int(src(i, 1) * 1.1)\nNext i\nws.Range("I2").Resize(UBound(dst, 1), 1).Value = dst   \' ← 書くのも1回</pre><p>ポイントが3つあります。<b>①</b> 複数セルの <code>.Value</code> は常に<b>2次元配列</b>（<code>src(行, 列)</code>）になります。1列でも <code>src(i, 1)</code> です。<b>②</b> 添字は <b>1始まり</b>。<b>③</b> 書き戻す範囲のサイズを <code>Resize</code> で配列とぴったり合わせます。</p>',
  hints:['<code>src = ws.Range("H2:H41").Value</code> だけで40個ぜんぶ読み込めます。','<code>ReDim dst(1 To 40, 1 To 1)</code> のように、書き戻す側も2次元で作ります。','<code>ws.Range("I2").Resize(40, 1).Value = dst</code>'],
  starter:'Option Explicit\n\nSub 練習13()\n    Dim ws As Worksheet\n    Dim src As Variant, dst() As Variant\n    Dim i As Long\n    Set ws = ThisWorkbook.Worksheets("売上明細")\n\n    ws.Range("I1").Value = "税込金額"\n\n\nEnd Sub',
  solution:'Option Explicit\n\nSub 練習13()\n    Dim ws As Worksheet\n    Dim src As Variant, dst() As Variant\n    Dim i As Long\n    Set ws = ThisWorkbook.Worksheets("売上明細")\n\n    ws.Range("I1").Value = "税込金額"\n\n    src = ws.Range("H2:H41").Value\n    ReDim dst(1 To UBound(src, 1), 1 To 1)\n\n    For i = 1 To UBound(src, 1)\n        dst(i, 1) = Int(src(i, 1) * 1.1)\n    Next i\n\n    ws.Range("I2").Resize(UBound(dst, 1), 1).Value = dst\nEnd Sub',
  check:function(ctx){
    if (txt(cell(ctx,'売上明細',1,9)) !== '税込金額') return fail('I1 に「税込金額」と書いてください。');
    var rows = baseSales(ctx);
    for (var i=0;i<40;i++){
      var want = Math.floor(rows[i].amt * 1.1);
      var got = cell(ctx,'売上明細',i+2,9);
      if (!eqNum(got, want)) return fail((i+2) + '行目が「' + txt(got) + '」です。Int(' + rows[i].amt + ' * 1.1) = ' + want + ' になるはずです。');
    }
    if (!/\.value/i.test(ctx.code) || !/(resize|ubound)/i.test(ctx.code))
      return fail('値は合っていますが、配列を使った形になっていません。Range の .Value をまるごと変数に入れ、UBound と Resize で書き戻してください。');
    return pass('読み書き各1回で40行を処理できました。実務ではこの差が数十倍になります。');
  }
},
{
  id:'l14', level:4, title:'文字列を整える', sheet:'要整形データ',
  goal:'「要整形データ」の <b>D列</b>（2〜10行目）に、C列の契約金額を<b>数値として</b>入れてください。全角数字・全角カンマ・半角カンマ・前後の空白が混ざっています。',
  body:'<p>人間が入力したデータは必ず汚れています。整形の道具は3つ覚えれば足ります。</p><ul><li><code>StrConv(s, vbNarrow)</code> … 全角を半角に（<code>１２０</code> → <code>120</code>、全角カンマ <code>，</code> → <code>,</code>）</li><li><code>Replace(s, ",", "")</code> … 文字を消す・置き換える</li><li><code>Trim(s)</code> … 前後の空白を落とす</li></ul><p>順番が大事です。<b>先に半角へ寄せてから</b>カンマを消すと、全角カンマも一緒に片付きます。逆順だと取りこぼします。</p><p>最後に <code>IsNumeric</code> で数値になるか確かめてから <code>CDbl</code> で変換すると、想定外の値で落ちません。</p>',
  hints:['<code>s = StrConv(ws.Cells(i, 3).Value, vbNarrow)</code> から始めます。','<code>s = Replace(s, ",", "")</code> → <code>s = Trim(s)</code>','<code>If IsNumeric(s) Then ws.Cells(i, 4).Value = CDbl(s)</code>'],
  starter:'Option Explicit\n\nSub 練習14()\n    Dim ws As Worksheet\n    Dim i As Long, s As String\n    Set ws = ThisWorkbook.Worksheets("要整形データ")\n\n    For i = 2 To 10\n\n    Next i\nEnd Sub',
  solution:'Option Explicit\n\nSub 練習14()\n    Dim ws As Worksheet\n    Dim i As Long, s As String\n    Set ws = ThisWorkbook.Worksheets("要整形データ")\n\n    For i = 2 To 10\n        s = StrConv(ws.Cells(i, 3).Value, vbNarrow)\n        s = Replace(s, ",", "")\n        s = Trim(s)\n\n        If IsNumeric(s) Then\n            ws.Cells(i, 4).Value = CDbl(s)\n        End If\n    Next i\nEnd Sub',
  check:function(ctx){
    var want = [120000, 85000, 240000, 56000, 330000, 1450000, 78000, 92000, 305000];
    for (var i=0;i<9;i++){
      var got = cell(ctx,'要整形データ',i+2,4);
      if (typeof got !== 'number') return fail((i+2) + '行目 D列が数値になっていません（今は「' + txt(got) + '」）。CDbl で数値に変換してください。');
      if (!eqNum(got, want[i])) return fail((i+2) + '行目が ' + txt(got) + ' です。正しくは ' + want[i].toLocaleString() + ' です。' + (i===0||i===4||i===6||i===8 ? ' 全角が残っていませんか？ StrConv を先にかけてください。' : ''));
    }
    return pass('9件すべて数値化できました。StrConv → Replace → Trim の順番が効いています。');
  }
},
{
  id:'l15', level:4, title:'Dictionary で集計する', sheet:'集計結果',
  goal:'担当者ごとの売上合計を求め、「集計結果」の <code>A列</code> に担当者名、<code>B列</code> に合計を、<b>1行目から</b>書き出してください。担当者は5人です。',
  body:'<p>Lv3 の支店別集計は、支店の一覧を自分で <code>Array</code> に書いていました。でも実務では「担当者が誰か」は事前に分かりません。</p><p><code>Dictionary</code> は「キーと値の組」を持つ入れ物です。集計の核はこの1行だけ。</p><pre class="mini">dic(key) = dic(key) + 金額</pre><p>存在チェックすら要りません。無いキーを読むと 0 が返るので、そのまま足せます。<b>ループは1周だけ</b>。二重ループが消えます。</p><p>取り出すときは <code>For Each k In dic.Keys</code>。<code>k</code> は <code>Variant</code> で宣言します。</p><p class="note"><b>実務での注意：</b><code>CreateObject("Scripting.Dictionary")</code> は <b>Windows 版 Excel 専用</b>です。Mac では使えないので <code>Collection</code> や配列で代用します。</p>',
  hints:['<code>Dim dic As Object</code> → <code>Set dic = CreateObject("Scripting.Dictionary")</code>','担当者は C列＝3列目、金額は H列＝8列目。','<code>Dim k As Variant</code> と <code>Dim r As Long</code> を用意して <code>For Each k In dic.Keys</code> で書き出します。'],
  starter:'Option Explicit\n\nSub 練習15()\n    Dim wsS As Worksheet, wsO As Worksheet\n    Dim dic As Object\n    Dim i As Long, r As Long\n    Dim k As Variant\n    Set wsS = ThisWorkbook.Worksheets("売上明細")\n    Set wsO = ThisWorkbook.Worksheets("集計結果")\n    Set dic = CreateObject("Scripting.Dictionary")\n\n    For i = 2 To 41\n\n    Next i\n\n    r = 1\n    For Each k In dic.Keys\n\n    Next k\nEnd Sub',
  solution:'Option Explicit\n\nSub 練習15()\n    Dim wsS As Worksheet, wsO As Worksheet\n    Dim dic As Object\n    Dim i As Long, r As Long\n    Dim k As Variant\n    Set wsS = ThisWorkbook.Worksheets("売上明細")\n    Set wsO = ThisWorkbook.Worksheets("集計結果")\n    Set dic = CreateObject("Scripting.Dictionary")\n\n    For i = 2 To 41\n        dic(wsS.Cells(i, 3).Value) = dic(wsS.Cells(i, 3).Value) + wsS.Cells(i, 8).Value\n    Next i\n\n    r = 1\n    For Each k In dic.Keys\n        wsO.Cells(r, 1).Value = k\n        wsO.Cells(r, 2).Value = dic(k)\n        r = r + 1\n    Next k\nEnd Sub',
  check:function(ctx){
    var want = {};
    baseSales(ctx).forEach(function(r){ want[r.staff] = (want[r.staff] || 0) + r.amt; });
    var names = Object.keys(want);
    var s = ctx.wb.byName('集計結果'), seen = {};
    for (var r=1;r<=12;r++){
      var n = txt(s.getV(r,1));
      if (!n) continue;
      if (want[n] === undefined) return fail(r + '行目の「' + n + '」は担当者一覧にありません。');
      if (seen[n]) return fail('「' + n + '」が2回書き出されています。');
      seen[n] = true;
      if (!eqNum(s.getV(r,2), want[n])) return fail('「' + n + '」の合計が「' + txt(s.getV(r,2)) + '」です。正しくは ' + want[n].toLocaleString() + ' です。');
    }
    var miss = names.filter(function(n){ return !seen[n]; });
    if (miss.length) return fail('「' + miss.join('」「') + '」が書き出されていません（担当者は ' + names.length + ' 人います）。');
    return pass(names.length + '人分を1周のループで集計できました。');
  }
},
{
  id:'l16', level:5, title:'総合：クロス集計表', sheet:'集計結果',
  goal:'「集計結果」に <b>支店 × カテゴリ</b> のクロス集計表を作ってください。<br><code>A1</code>＝<b>支店</b>、<code>B1</code> から右へカテゴリ名、<code>A2</code> から下へ支店名、交差するセルに売上合計。さらに <b>最終列に「合計」</b>（その支店の総合計）を付けてください。<br>カテゴリは 文具・用紙・トナー・ファイル・備品 の5つ、支店は 東京・大阪・名古屋 です。',
  body:'<p>Lv1〜4 の道具を全部使います。ループ、条件分岐、配列、書式。実務の集計マクロはだいたいこの形です。</p><p><b>組み立て方</b></p><ol><li>見出しを書く（A1＝「支店」、B1〜F1＝カテゴリ、G1＝「合計」）</li><li>支店×カテゴリの2次元配列を <code>ReDim</code> で用意する</li><li>明細を<b>1周だけ</b>回して、該当するマスに足し込む</li><li>配列をシートに書き出し、行ごとの合計を出す</li></ol><p>支店名・カテゴリ名から「配列の何番目か」を求める小さな関数を <code>Function</code> で作っておくと、本体がすっきりします。<code>Function</code> は <b>自分の名前に値を代入</b>することで結果を返します。</p><pre class="mini">Function IndexOf(arr As Variant, v As String) As Long\n    Dim i As Long\n    For i = 0 To UBound(arr)\n        If arr(i) = v Then\n            IndexOf = i + 1\n            Exit Function\n        End If\n    Next i\nEnd Function</pre><p>合計列は <code>WorksheetFunction.Sum</code> を使ってもループで足してもかまいません。</p>',
  hints:['<code>ReDim grid(1 To 3, 1 To 5)</code>（支店3 × カテゴリ5）','支店は B列＝2列目、カテゴリは E列＝5列目、金額は H列＝8列目。','書き出す位置は、支店 b 行目なら <code>Cells(b + 1, ...)</code>、カテゴリ c 列目なら <code>Cells(..., c + 1)</code>。','合計列は7列目（G列）です。'],
  starter:'Option Explicit\n\nSub 練習16()\n    Dim wsS As Worksheet, wsO As Worksheet\n    Dim branches As Variant, cats As Variant\n    Dim grid() As Currency\n    Dim i As Long, b As Long, c As Long\n    Set wsS = ThisWorkbook.Worksheets("売上明細")\n    Set wsO = ThisWorkbook.Worksheets("集計結果")\n\n    branches = Array("東京", "大阪", "名古屋")\n    cats = Array("文具", "用紙", "トナー", "ファイル", "備品")\n    ReDim grid(1 To 3, 1 To 5)\n\n    \' 1) 見出し\n\n    \' 2) 明細を1周して足し込む\n\n    \' 3) 書き出し ＋ 合計列\n\nEnd Sub\n\nFunction IndexOf(arr As Variant, v As String) As Long\n    Dim i As Long\n    For i = 0 To UBound(arr)\n        If arr(i) = v Then\n            IndexOf = i + 1\n            Exit Function\n        End If\n    Next i\nEnd Function',
  solution:'Option Explicit\n\nSub 練習16()\n    Dim wsS As Worksheet, wsO As Worksheet\n    Dim branches As Variant, cats As Variant\n    Dim grid() As Currency\n    Dim i As Long, b As Long, c As Long\n    Dim total As Currency\n    Set wsS = ThisWorkbook.Worksheets("売上明細")\n    Set wsO = ThisWorkbook.Worksheets("集計結果")\n\n    branches = Array("東京", "大阪", "名古屋")\n    cats = Array("文具", "用紙", "トナー", "ファイル", "備品")\n    ReDim grid(1 To 3, 1 To 5)\n\n    wsO.Range("A1").Value = "支店"\n    For c = 0 To UBound(cats)\n        wsO.Cells(1, c + 2).Value = cats(c)\n    Next c\n    wsO.Cells(1, UBound(cats) + 3).Value = "合計"\n\n    For i = 2 To 41\n        b = IndexOf(branches, wsS.Cells(i, 2).Value)\n        c = IndexOf(cats, wsS.Cells(i, 5).Value)\n        If b > 0 And c > 0 Then\n            grid(b, c) = grid(b, c) + wsS.Cells(i, 8).Value\n        End If\n    Next i\n\n    For b = 1 To 3\n        wsO.Cells(b + 1, 1).Value = branches(b - 1)\n        total = 0\n        For c = 1 To 5\n            wsO.Cells(b + 1, c + 1).Value = grid(b, c)\n            total = total + grid(b, c)\n        Next c\n        wsO.Cells(b + 1, 7).Value = total\n    Next b\n\n    wsO.Range("A1:G1").Font.Bold = True\n    wsO.Range("B2:G4").NumberFormat = "#,##0"\nEnd Sub\n\nFunction IndexOf(arr As Variant, v As String) As Long\n    Dim i As Long\n    For i = 0 To UBound(arr)\n        If arr(i) = v Then\n            IndexOf = i + 1\n            Exit Function\n        End If\n    Next i\nEnd Function',
  check:function(ctx){
    var cats = ['文具','用紙','トナー','ファイル','備品'];
    var grid = {}; BRANCHES.forEach(function(b){ grid[b] = {}; cats.forEach(function(c){ grid[b][c] = 0; }); });
    baseSales(ctx).forEach(function(r){ grid[r.branch][r.cat] += r.amt; });
    var s = ctx.wb.byName('集計結果');
    if (txt(s.getV(1,1)) !== '支店') return fail('A1 に「支店」と書いてください。');
    var colOf = {};
    for (var c=2;c<=6;c++){
      var name = txt(s.getV(1,c));
      if (cats.indexOf(name) < 0) return fail(global.VBARuntime.colLetter(c) + '1 が「' + (name||'空') + '」です。B1〜F1 にカテゴリ名（' + cats.join('・') + '）を並べてください。');
      if (colOf[name]) return fail('カテゴリ「' + name + '」が2回出ています。');
      colOf[name] = c;
    }
    if (txt(s.getV(1,7)) !== '合計') return fail('G1 に「合計」と書いてください。');
    var rowOf = {};
    for (var r=2;r<=4;r++){
      var bn = txt(s.getV(r,1));
      if (BRANCHES.indexOf(bn) < 0) return fail('A' + r + ' が「' + (bn||'空') + '」です。A2〜A4 に支店名を並べてください。');
      rowOf[bn] = r;
    }
    for (var bi=0;bi<BRANCHES.length;bi++){
      var b = BRANCHES[bi], rr = rowOf[b];
      if (!rr) return fail('支店「' + b + '」の行がありません。');
      var sum = 0;
      for (var ci=0;ci<cats.length;ci++){
        var cat = cats[ci], want = grid[b][cat], got = s.getV(rr, colOf[cat]);
        var blank = (got === '' || got === undefined || got === null || (got && got.__empty));
        if (!eqNum(got, want) && !(want === 0 && blank))
          return fail(b + ' × ' + cat + ' が「' + txt(got) + '」です。正しくは ' + want.toLocaleString() + ' です。');
        sum += want;
      }
      if (!eqNum(s.getV(rr,7), sum)) return fail(b + ' の合計列が「' + txt(s.getV(rr,7)) + '」です。正しくは ' + sum.toLocaleString() + ' です。');
    }
    return pass('クロス集計表が完成しました。これで実務のマクロが書けます。');
  }
}
];

var LEVELS = [
  { n:1, name:'基礎', desc:'VBE に慣れる' },
  { n:2, name:'制御', desc:'条件分岐とループ' },
  { n:3, name:'実務', desc:'現場で使う操作' },
  { n:4, name:'応用', desc:'速く、壊れないコードへ' },
  { n:5, name:'総合', desc:'ぜんぶ使う' }
];

global.VBALessons = { lessons: LESSONS, levels: LEVELS, buildWorkbook: buildWorkbook, branches: BRANCHES, products: PRODUCTS };
})(window);
