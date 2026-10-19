/* 読みもの：VBA の仕組み（図解） */
window.VBAGuide = { html: `
<div class="guide">
<div class="eyebrow">読みもの</div>
<h2>VBAの仕組み</h2>

<p>文法を覚える前に、VBA が何をしているのかを3枚の図で押さえておきます。ここが分かると、書き方の理由とエラーの意味が同時に見えるようになります。</p>

<h3>1. コードはExcelの外で動いている</h3>

<p>VBA のコードはセルの中にいるわけではありません。Excel の隣に座っている別の存在で、何をするにも「Excel に頼む」という形をとります。</p>

<div class="figwrap"><svg class="fig" viewBox="0 0 680 324" role="img" aria-label="VBA側とExcel側が窓口越しに往復でやり取りする図">
<defs><marker id="fa1" viewBox="0 0 10 10" refX="8" refY="5" markerWidth="6" markerHeight="6" orient="auto-start-reverse"><path d="M2 1L8 5L2 9" fill="none" stroke="var(--ink2)" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></marker></defs>
<text class="th" x="40" y="50">VBA（マクロ）</text>
<rect class="dash" x="40" y="62" width="230" height="165" rx="4"/>
<g class="g-1"><rect x="60" y="89" width="190" height="46" rx="4"/><text class="th" x="155" y="109" text-anchor="middle">コード</text><text class="ts" x="155" y="126" text-anchor="middle">Sub 〜 End Sub</text></g>
<g class="g-1"><rect x="60" y="154" width="190" height="46" rx="4"/><text class="th" x="155" y="174" text-anchor="middle">変数</text><text class="ts" x="155" y="191" text-anchor="middle">price / qty / amount</text></g>
<text class="th" x="410" y="50">Excel（ブック）</text>
<rect class="dash" x="410" y="62" width="230" height="165" rx="4"/>
<text class="ts" x="525" y="89" text-anchor="middle">ワークシート</text>
<rect class="box" x="437" y="99" width="44" height="22" rx="2"/><rect class="box" x="481" y="99" width="44" height="22" rx="2"/><rect class="box" x="525" y="99" width="44" height="22" rx="2"/><rect class="box" x="569" y="99" width="44" height="22" rx="2"/>
<rect class="box" x="437" y="121" width="44" height="22" rx="2"/><rect class="box" x="525" y="121" width="44" height="22" rx="2"/><rect class="box" x="569" y="121" width="44" height="22" rx="2"/>
<g class="g-2"><rect x="481" y="121" width="44" height="22" rx="2"/><text class="ts" x="503" y="136" text-anchor="middle">F2</text></g>
<rect class="box" x="437" y="143" width="44" height="22" rx="2"/><rect class="box" x="481" y="143" width="44" height="22" rx="2"/><rect class="box" x="525" y="143" width="44" height="22" rx="2"/><rect class="box" x="569" y="143" width="44" height="22" rx="2"/>
<rect class="box" x="437" y="165" width="44" height="22" rx="2"/><rect class="box" x="481" y="165" width="44" height="22" rx="2"/><rect class="box" x="525" y="165" width="44" height="22" rx="2"/><rect class="box" x="569" y="165" width="44" height="22" rx="2"/>
<text class="ts" x="525" y="209" text-anchor="middle">値・数式・書式</text>
<g class="g-3"><rect x="282" y="90" width="116" height="26" rx="4"/><text class="ts" x="340" y="107" text-anchor="middle">F2 の値は？</text></g>
<line class="arr" x1="270" y1="124" x2="410" y2="124" marker-end="url(#fa1)"/>
<line class="arr" x1="410" y1="189" x2="270" y2="189" marker-end="url(#fa1)"/>
<g class="g-3"><rect x="282" y="200" width="116" height="26" rx="4"/><text class="ts" x="340" y="217" text-anchor="middle">780 が返る</text></g>
<g class="g-2"><rect x="40" y="254" width="600" height="50" rx="4"/><text class="th" x="340" y="276" text-anchor="middle">セルに1回触るたび、この往復が1回起きる</text><text class="ts" x="340" y="294" text-anchor="middle">これが VBA の速さを決める</text></g>
</svg></div>

<p><code>ws.Range("F2").Value</code> は「F2の中身を見る」のではなく、「F2の中身を教えてくれと頼んで、返ってきた答えを受け取る」という往復です。</p>

<p>シート指定を省いてはいけない理由もここにあります。相手を指定しないまま注文を出すと、<b>今たまたま開いているシート</b>に届いてしまうからです。</p>

<pre class="mini">ws.Range("F2").Value      ' 宛先がはっきりしている
Range("F2").Value         ' 今開いているシートに届く（事故のもと）</pre>

<h3>2. ドットは、階層を降りる道順</h3>

<p>VBA の命令は必ず何かのオブジェクトに向かって出されます。そのオブジェクトは入れ子になっていて、コードに出てくるドットは、その階層を1段ずつ内側に降りる動きそのものです。</p>

<div class="figwrap"><svg class="fig" viewBox="0 0 680 370" role="img" aria-label="Application、Workbook、Worksheet、Range、プロパティという入れ子の階層図">
<rect class="box" x="40" y="35" width="600" height="38" rx="4"/>
<text class="t" x="340" y="59" text-anchor="middle">ThisWorkbook.Worksheets("売上明細").Range("F2").Value</text>
<rect class="dash" x="40" y="90" width="600" height="260" rx="4"/>
<text class="th" x="52" y="108">Application — Excel そのもの</text>
<rect class="dash" x="75" y="120" width="530" height="205" rx="4"/>
<text class="th" x="87" y="138">Workbook — ThisWorkbook</text>
<rect class="dash" x="110" y="150" width="460" height="150" rx="4"/>
<text class="th" x="122" y="168">Worksheet — Worksheets("売上明細")</text>
<rect class="dash" x="145" y="180" width="390" height="95" rx="4"/>
<text class="th" x="157" y="198">Range — Range("F2")</text>
<g class="g-1"><rect x="155" y="210" width="110" height="44" rx="4"/><text class="th" x="210" y="229" text-anchor="middle">Value</text><text class="ts" x="210" y="246" text-anchor="middle">中身の値</text></g>
<g class="g-1"><rect x="285" y="210" width="110" height="44" rx="4"/><text class="th" x="340" y="229" text-anchor="middle">Font</text><text class="ts" x="340" y="246" text-anchor="middle">文字の書式</text></g>
<g class="g-1"><rect x="415" y="210" width="110" height="44" rx="4"/><text class="th" x="470" y="229" text-anchor="middle">Interior</text><text class="ts" x="470" y="246" text-anchor="middle">塗りつぶし</text></g>
</svg></div>

<p>長いコードに見えても、やっていることは住所をたどっているだけです。ブックの中の、シートの中の、セルの中の、値。特別な文法ではありません。</p>

<p>この構造が分かると、エラーの読み方が変わります。</p>

<ul>
<li><code>オブジェクトが必要です</code> — 住所の途中が空っぽ。<code>Set</code> の書き忘れがほとんどです</li>
<li><code>インデックスが有効範囲にありません</code> — 指定した住所にそんなものが無い。シート名のタイプミスが典型です</li>
<li><code>このオブジェクトに◯◯はありません</code> — その階層に存在しないプロパティを呼んでいる。<code>Worksheet</code> に <code>.Value</code> は無い、など</li>
</ul>

<h3>3. 往復の回数が速さを決める</h3>

<p>同じ答えを出すコードでも、Excel に話しかける回数は書き方で決まります。</p>

<div class="figwrap"><svg class="fig" viewBox="0 0 680 335" role="img" aria-label="1セルずつ触ると40往復、配列でまとめて読むと1往復で済むことの比較図">
<g class="g-2"><rect x="40" y="75" width="185" height="52" rx="4"/><text class="th" x="132" y="100" text-anchor="middle">1セルずつ触る</text><text class="ts" x="132" y="117" text-anchor="middle">ws.Cells(i, 8).Value</text></g>
<g class="g-2"><rect x="255" y="82" width="10" height="10" rx="2"/><rect x="269" y="82" width="10" height="10" rx="2"/><rect x="283" y="82" width="10" height="10" rx="2"/><rect x="297" y="82" width="10" height="10" rx="2"/><rect x="311" y="82" width="10" height="10" rx="2"/><rect x="325" y="82" width="10" height="10" rx="2"/><rect x="339" y="82" width="10" height="10" rx="2"/><rect x="353" y="82" width="10" height="10" rx="2"/><rect x="367" y="82" width="10" height="10" rx="2"/><rect x="381" y="82" width="10" height="10" rx="2"/><rect x="395" y="82" width="10" height="10" rx="2"/><rect x="409" y="82" width="10" height="10" rx="2"/><rect x="423" y="82" width="10" height="10" rx="2"/><rect x="437" y="82" width="10" height="10" rx="2"/><rect x="451" y="82" width="10" height="10" rx="2"/><rect x="465" y="82" width="10" height="10" rx="2"/><rect x="479" y="82" width="10" height="10" rx="2"/><rect x="493" y="82" width="10" height="10" rx="2"/><rect x="507" y="82" width="10" height="10" rx="2"/><rect x="521" y="82" width="10" height="10" rx="2"/><rect x="255" y="98" width="10" height="10" rx="2"/><rect x="269" y="98" width="10" height="10" rx="2"/><rect x="283" y="98" width="10" height="10" rx="2"/><rect x="297" y="98" width="10" height="10" rx="2"/><rect x="311" y="98" width="10" height="10" rx="2"/><rect x="325" y="98" width="10" height="10" rx="2"/><rect x="339" y="98" width="10" height="10" rx="2"/><rect x="353" y="98" width="10" height="10" rx="2"/><rect x="367" y="98" width="10" height="10" rx="2"/><rect x="381" y="98" width="10" height="10" rx="2"/><rect x="395" y="98" width="10" height="10" rx="2"/><rect x="409" y="98" width="10" height="10" rx="2"/><rect x="423" y="98" width="10" height="10" rx="2"/><rect x="437" y="98" width="10" height="10" rx="2"/><rect x="451" y="98" width="10" height="10" rx="2"/><rect x="465" y="98" width="10" height="10" rx="2"/><rect x="479" y="98" width="10" height="10" rx="2"/><rect x="493" y="98" width="10" height="10" rx="2"/><rect x="507" y="98" width="10" height="10" rx="2"/><rect x="521" y="98" width="10" height="10" rx="2"/></g>
<text class="th" x="555" y="99">40 往復</text>
<g class="g-1"><rect x="40" y="185" width="185" height="52" rx="4"/><text class="th" x="132" y="210" text-anchor="middle">配列でまとめて</text><text class="ts" x="132" y="227" text-anchor="middle">Range("H2:H41").Value</text></g>
<g class="g-1"><rect x="255" y="206" width="10" height="10" rx="2"/></g>
<text class="th" x="555" y="215">1 往復</text>
<g class="g-3"><rect x="40" y="265" width="600" height="50" rx="4"/><text class="th" x="340" y="287" text-anchor="middle">答えは同じ。Excel に話しかけた回数だけが違う</text><text class="ts" x="340" y="305" text-anchor="middle">実行すると出る「Excel とのやり取り ◯回」がこの数</text></g>
</svg></div>

<p>実務のマクロが「動くけれど遅い」から「実用的」に変わる分かれ目がここです。レベル4の「配列で高速化する」で、この数字が実際に 40 から 3 まで落ちるのを確かめられます。</p>

<p class="note"><b>補足。</b><code>Dim</code> した変数は VBA 側にあるので、変数同士の計算は何度やっても往復ゼロです。<code>amount = price * qty</code> にコストが無いのはそのため。だから「まとめて読む → VBA側で計算する → まとめて書く」という順番になります。</p>

<div class="helpers"><button class="mini-btn" id="guideStart">最初のレッスンを始める</button></div>
</div>
`};
