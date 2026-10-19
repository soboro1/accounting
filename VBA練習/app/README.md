# VBA道場（アプリのソース）

公開先: https://claude.ai/artifact/4dC7uyYMHqKxjULAucY4Em

| ファイル | 中身 |
|---|---|
| `index.html` | 画面・スタイル・アプリ本体 |
| `vba.js` | VBA の字句解析・構文解析（トークナイザ＋パーサ） |
| `runtime.js` | 擬似 Excel ワークブックと評価器（Range / Cells / Font / 組み込み関数） |
| `lessons.js` | サンプルデータと全16問（課題文・ヒント・解答例・採点ロジック） |
| `guide.js` | 読みもの「VBAの仕組み」の本文とSVG図3枚 |

## 直したいとき

`lessons.js` の `LESSONS` 配列を編集すれば問題を足せます。1問の形:

```js
{
  id:'l17', level:5, title:'タイトル', sheet:'売上明細',
  goal:'やること（HTML可）',
  body:'解説（HTML可）',
  hints:['ヒント1','ヒント2'],
  starter:'Sub 練習17()\n\nEnd Sub',
  solution:'...',
  check:function(ctx){ /* ctx.wb=実行後 / ctx.base=実行前 / ctx.out=MsgBox / ctx.code=書いたコード */
    return { ok:true, msg:'正解です。' };
  }
}
```

編集後に、全問の解答が採点を通るか確認:

```bash
node -e "global.window={};['vba','runtime','lessons'].forEach(f=>require('./'+f+'.js'));var RT=window.VBARuntime,L=window.VBALessons;L.lessons.forEach(l=>{var wb=L.buildWorkbook(),b=L.buildWorkbook(),r=RT.run(l.solution,wb,{});console.log((r.ok&&l.check({wb,base:b,out:r.out,code:l.solution}).ok?'✓ ':'✗ ')+l.id)})"
```

再公開は Claude Code から `index.html` を同じ URL に publish し直します。
