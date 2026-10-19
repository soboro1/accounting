/* ================================================================
   擬似 Excel ワークブック ＋ 評価器
   ================================================================ */
(function (global) {
'use strict';
var I = global.VBAInternals;
var EMPTY = I.EMPTY, NOTHING = I.NOTHING, VbaArray = I.VbaArray, VbaError = I.VbaError,
    ErrValue = I.ErrValue, MAXROW = I.MAXROW, MAXCOL = I.MAXCOL;

/* ---------- A1 形式 ---------- */
function colLetter(c){ var s = ''; while (c > 0){ var m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = (c - m - 1) / 26; } return s; }
function letterCol(s){ var c = 0; s = s.toUpperCase(); for (var i=0;i<s.length;i++) c = c * 26 + (s.charCodeAt(i) - 64); return c; }
function parseCellRef(s){
  var m = /^\$?([A-Za-z]{1,3})\$?([0-9]{1,7})$/.exec(s.trim());
  if (!m) return null;
  return { r: parseInt(m[2], 10), c: letterCol(m[1]) };
}
function parseAddress(sheet, addr){
  addr = String(addr).trim();
  var parts = addr.split(':');
  if (parts.length === 1){
    var a = parseCellRef(parts[0]);
    if (!a) throw new VbaError(1004, 'アドレス「' + addr + '」を解釈できません');
    return mkRange(sheet, a.r, a.c, a.r, a.c);
  }
  var colOnly = /^\$?[A-Za-z]{1,3}$/;
  if (colOnly.test(parts[0]) && colOnly.test(parts[1])){
    var c1 = letterCol(parts[0].replace('$','')), c2 = letterCol(parts[1].replace('$',''));
    return mkRange(sheet, 1, Math.min(c1,c2), MAXROW, Math.max(c1,c2));
  }
  if (/^\$?[0-9]+$/.test(parts[0])){
    var r1 = parseInt(parts[0].replace('$',''),10), r2 = parseInt(parts[1].replace('$',''),10);
    return mkRange(sheet, Math.min(r1,r2), 1, Math.max(r1,r2), MAXCOL);
  }
  var p = parseCellRef(parts[0]), q = parseCellRef(parts[1]);
  if (!p || !q) throw new VbaError(1004, 'アドレス「' + addr + '」を解釈できません');
  return mkRange(sheet, Math.min(p.r,q.r), Math.min(p.c,q.c), Math.max(p.r,q.r), Math.max(p.c,q.c));
}

/* ---------- シート ---------- */
var ACCESS = { n: 0 };
function Sheet(name){ this.name = name; this.cells = new Map(); this.colWidth = new Map(); }
Sheet.prototype.key = function(r,c){ return r + ':' + c; };
Sheet.prototype.cell = function(r,c,create){
  var k = this.key(r,c), x = this.cells.get(k);
  if (!x && create){ x = { v: EMPTY }; this.cells.set(k, x); }
  return x;
};
Sheet.prototype.getV = function(r,c){ var x = this.cells.get(this.key(r,c)); return x ? x.v : EMPTY; };
Sheet.prototype.setV = function(r,c,v){
  if (r < 1 || c < 1 || r > MAXROW || c > MAXCOL) throw new VbaError(1004, 'セルの位置が範囲外です');
  var x = this.cell(r,c,true); x.v = v;
};
Sheet.prototype.isBlank = function(r,c){ var v = this.getV(r,c); return v === EMPTY || v === '' || v === undefined || v === null; };
Sheet.prototype.bounds = function(){
  var maxR = 0, maxC = 0;
  this.cells.forEach(function(x, k){
    if (x.v === EMPTY && !x.bold && !x.fill && !x.fcolor) return;
    var p = k.split(':'), r = +p[0], c = +p[1];
    if (r > maxR) maxR = r; if (c > maxC) maxC = c;
  });
  return { rows: maxR, cols: maxC };
};
Sheet.prototype.clone = function(){
  var s = new Sheet(this.name);
  this.cells.forEach(function(x,k){ s.cells.set(k, Object.assign({}, x)); });
  this.colWidth.forEach(function(w,k){ s.colWidth.set(k,w); });
  return s;
};

function Workbook(sheets){ this.sheets = sheets; this.active = sheets[0]; this.name = 'サンプルデータ.xlsm'; }
Workbook.prototype.byName = function(n){
  for (var i=0;i<this.sheets.length;i++) if (this.sheets[i].name.toLowerCase() === String(n).toLowerCase()) return this.sheets[i];
  return null;
};
Workbook.prototype.clone = function(){
  var s = this.sheets.map(function(x){ return x.clone(); });
  var wb = new Workbook(s);
  wb.active = s[0];
  return wb;
};

/* ---------- ホストオブジェクト ---------- */
function mkRange(sheet,r1,c1,r2,c2){ return { _t:'Range', sheet:sheet, r1:r1, c1:c1, r2:r2, c2:c2 }; }
function mkSheetObj(s){ return { _t:'Sheet', sheet:s }; }

/* ---------- 型変換 ---------- */
var DAY = 86400000, EPOCH = Date.UTC(1899,11,30);
function dateToSerial(d){ return (Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), d.getHours(), d.getMinutes(), d.getSeconds()) - EPOCH) / DAY; }
function serialToDate(n){ var ms = EPOCH + Math.round(n * DAY); var d = new Date(ms); return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate(), d.getUTCHours(), d.getUTCMinutes(), d.getUTCSeconds()); }
function isDate(v){ return v instanceof Date; }

function toNum(v){
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? -1 : 0;
  if (v === EMPTY || v === undefined || v === null) return 0;
  if (isDate(v)) return dateToSerial(v);
  if (typeof v === 'string'){
    var s = v.trim().replace(/,/g, '');
    if (s === '') return 0;
    var n = Number(s);
    if (isNaN(n)) throw new VbaError(13, '型が一致しません（「' + v + '」は数値になりません）');
    return n;
  }
  if (v && v._t === 'Range') return toNum(rangeValue(v));
  throw new VbaError(13, '型が一致しません');
}
function numStr(n){
  if (!isFinite(n)) return String(n);
  if (Number.isInteger(n)) return String(n);
  var s = n.toPrecision(15).replace(/0+$/,'').replace(/\.$/,'');
  return String(parseFloat(s));
}
function toStr(v){
  if (typeof v === 'string') return v;
  if (typeof v === 'number') return numStr(v);
  if (typeof v === 'boolean') return v ? 'True' : 'False';
  if (v === EMPTY || v === undefined || v === null) return '';
  if (v === NOTHING) return 'Nothing';
  if (isDate(v)) return fmtDate(v, v.getHours()||v.getMinutes() ? 'yyyy/mm/dd hh:mm:ss' : 'yyyy/mm/dd');
  if (v instanceof VbaArray) throw new VbaError(13, '配列をそのまま文字列にはできません');
  if (v && v._t === 'Range') return toStr(rangeValue(v));
  if (v && v._t === 'ErrorValue') return 'Error ' + v.number;
  throw new VbaError(13, '型が一致しません');
}
function toBool(v){
  if (typeof v === 'boolean') return v;
  if (v === EMPTY || v === undefined || v === null) return false;
  if (typeof v === 'string'){ if (v.toLowerCase()==='true') return true; if (v.toLowerCase()==='false') return false; return toNum(v) !== 0; }
  if (v && v._t === 'Range') return toBool(rangeValue(v));
  return toNum(v) !== 0;
}

/* ---------- 書式 ---------- */
function pad(n,w){ n = String(n); while (n.length < w) n = '0' + n; return n; }
var JP_DAY = ['日','月','火','水','木','金','土'];
function fmtDate(d, f){
  return f.replace(/yyyy|yy|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|nn|n|ss|s|aaa/g, function(t){
    switch(t){
      case 'yyyy': return d.getFullYear();
      case 'yy': return pad(d.getFullYear()%100,2);
      case 'mm': return pad(d.getMonth()+1,2);
      case 'm': return d.getMonth()+1;
      case 'dd': return pad(d.getDate(),2);
      case 'd': return d.getDate();
      case 'aaa': return JP_DAY[d.getDay()];
      case 'hh': return pad(d.getHours(),2);
      case 'h': return d.getHours();
      case 'nn': return pad(d.getMinutes(),2);
      case 'n': return d.getMinutes();
      case 'ss': return pad(d.getSeconds(),2);
      case 's': return d.getSeconds();
      default: return t;
    }
  });
}
function fmtNumber(n, f){
  var pre = '', post = '';
  f = f.replace(/"([^"]*)"/g, function(_, s){ pre += s; return ''; });
  var pct = f.indexOf('%') >= 0;
  if (pct){ n = n * 100; f = f.replace(/%/g,''); post = '%'; }
  if (f.indexOf('¥') >= 0){ pre = '¥' + pre; f = f.replace(/¥/g,''); }
  var neg = n < 0; n = Math.abs(n);
  var dot = f.indexOf('.');
  var dec = dot >= 0 ? (f.slice(dot+1).match(/[0#]/g) || []).length : 0;
  var group = f.indexOf(',') >= 0;
  var s = n.toFixed(dec);
  var parts = s.split('.');
  if (group) parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  s = parts.join('.');
  return (neg ? '-' : '') + pre + s + post;
}
function applyFormat(v, f){
  if (!f || f === 'General' || f === '@') return toStr(v);
  if (/[ymdhns]/.test(f) && !/[#0]/.test(f)){
    var d = isDate(v) ? v : (typeof v === 'number' ? serialToDate(v) : null);
    if (!d) return toStr(v);
    return fmtDate(d, f);
  }
  if (typeof v === 'string' && v.trim() === '') return v;
  var n; try { n = toNum(v); } catch(e){ return toStr(v); }
  return fmtNumber(n, f);
}
global.VBAFormat = applyFormat;

/* ---------- 範囲の値 ---------- */
function rangeValue(rg){
  ACCESS.n++;
  if (rg.r1 === rg.r2 && rg.c1 === rg.c2) return rg.sheet.getV(rg.r1, rg.c1);
  var rows = rg.r2 - rg.r1 + 1, cols = rg.c2 - rg.c1 + 1;
  if (rows > 200000 || cols > 1000) throw new VbaError(7, '範囲が大きすぎます');
  var arr = new VbaArray([{lb:1,ub:rows},{lb:1,ub:cols}]);
  for (var r=0;r<rows;r++) for (var c=0;c<cols;c++) arr.data[r*cols+c] = rg.sheet.getV(rg.r1+r, rg.c1+c);
  return arr;
}
function setRangeValue(rg, v){
  ACCESS.n++;
  var rows = rg.r2 - rg.r1 + 1, cols = rg.c2 - rg.c1 + 1;
  if (v instanceof VbaArray){
    var d = v.dims;
    var vr = d[0].ub - d[0].lb + 1, vc = d.length > 1 ? d[1].ub - d[1].lb + 1 : 1;
    for (var r=0;r<rows;r++) for (var c=0;c<cols;c++){
      if (r < vr && c < vc){
        var val = d.length > 1 ? v.data[r*vc+c] : (c === 0 ? v.data[r] : EMPTY);
        rg.sheet.setV(rg.r1+r, rg.c1+c, val);
      }
    }
    return;
  }
  if (rows * cols > 200000) throw new VbaError(7, '範囲が大きすぎます');
  for (var r2=0;r2<rows;r2++) for (var c2=0;c2<cols;c2++) rg.sheet.setV(rg.r1+r2, rg.c1+c2, v);
}
function rangeCells(rg){
  var out = [];
  var rows = Math.min(rg.r2, 5000), cols = Math.min(rg.c2, 200);
  for (var r=rg.r1;r<=rows;r++) for (var c=rg.c1;c<=cols;c++) out.push(mkRange(rg.sheet,r,c,r,c));
  return out;
}
function rangeFlatValues(rg){
  var out = [];
  for (var r=rg.r1;r<=Math.min(rg.r2, 200000);r++) for (var c=rg.c1;c<=Math.min(rg.c2,1000);c++) out.push(rg.sheet.getV(r,c));
  return out;
}
function endMove(rg, dir){
  var sh = rg.sheet, r = rg.r1, c = rg.c1;
  if (dir === 'up'){
    if (r === 1) return mkRange(sh,1,c,1,c);
    if (sh.isBlank(r-1,c)){ var rr = r-1; while (rr > 1 && sh.isBlank(rr,c)) rr--; return mkRange(sh,rr,c,rr,c); }
    var r2 = r; while (r2 > 1 && !sh.isBlank(r2-1,c)) r2--; return mkRange(sh,r2,c,r2,c);
  }
  if (dir === 'down'){
    var b = sh.bounds().rows;
    if (sh.isBlank(r+1,c)){ var rd = r+1; while (rd < MAXROW && sh.isBlank(rd,c) && rd <= b) rd++; return mkRange(sh,Math.min(rd, MAXROW),c,Math.min(rd,MAXROW),c); }
    var r3 = r; while (r3 < MAXROW && !sh.isBlank(r3+1,c)) r3++; return mkRange(sh,r3,c,r3,c);
  }
  if (dir === 'left'){
    if (c === 1) return mkRange(sh,r,1,r,1);
    if (sh.isBlank(r,c-1)){ var cc = c-1; while (cc > 1 && sh.isBlank(r,cc)) cc--; return mkRange(sh,r,cc,r,cc); }
    var c2 = c; while (c2 > 1 && !sh.isBlank(r,c2-1)) c2--; return mkRange(sh,r,c2,r,c2);
  }
  var cb = sh.bounds().cols;
  if (sh.isBlank(r,c+1)){ var cd = c+1; while (cd <= cb && sh.isBlank(r,cd)) cd++; return mkRange(sh,r,Math.min(cd,MAXCOL),r,Math.min(cd,MAXCOL)); }
  var c3 = c; while (c3 < MAXCOL && !sh.isBlank(r,c3+1)) c3++; return mkRange(sh,r,c3,r,c3);
}
function currentRegion(rg){
  var sh = rg.sheet, r1 = rg.r1, r2 = rg.r1, c1 = rg.c1, c2 = rg.c1;
  function rowUsed(r){ for (var c=Math.max(1,c1-1); c<=c2+1; c++) if (!sh.isBlank(r,c)) return true; return false; }
  function colUsed(c){ for (var r=Math.max(1,r1-1); r<=r2+1; r++) if (!sh.isBlank(r,c)) return true; return false; }
  var guard = 0;
  var changed = true;
  while (changed && guard++ < 5000){
    changed = false;
    while (r1 > 1 && rowUsed(r1-1)){ r1--; changed = true; }
    while (r2 < MAXROW && rowUsed(r2+1)){ r2++; changed = true; if (r2 > 200000) break; }
    while (c1 > 1 && colUsed(c1-1)){ c1--; changed = true; }
    while (c2 < MAXCOL && colUsed(c2+1)){ c2++; changed = true; if (c2 > 500) break; }
  }
  return mkRange(sh, r1, c1, r2, c2);
}

/* ---------- 定数 ---------- */
var CONSTS = {
  vbcrlf:'\r\n', vblf:'\n', vbcr:'\r', vbtab:'\t', vbnullstring:'', vbnewline:'\n',
  vbok:1, vbcancel:2, vbyes:6, vbno:7, vbabort:3, vbretry:4, vbignore:5,
  vbokonly:0, vbokcancel:1, vbyesno:4, vbyesnocancel:3, vbretrycancel:5,
  vbcritical:16, vbquestion:32, vbexclamation:48, vbinformation:64,
  vbnarrow:8, vbwide:4, vbuppercase:1, vblowercase:2, vbpropercase:3, vbkatakana:16, vbhiragana:32,
  xlup:'up', xldown:'down', xltoleft:'left', xltoright:'right',
  xlyes:1, xlno:2, xlascending:1, xldescending:2,
  xlcalculationmanual:-4135, xlcalculationautomatic:-4105,
  xlcelltypevisible:12, xltypepdf:0, xlqualitystandard:0,
  xllandscape:2, xlportrait:1, xlcontinuous:1, xlthin:2, xlcenter:-4108, xlleft:-4131, xlright:-4152,
  xlpart:2, xlwhole:1, xlvalues:-4163
};

/* ---------- インタプリタ ---------- */
var SIG_EXIT_FOR = { sig:'for' }, SIG_EXIT_DO = { sig:'do' }, SIG_EXIT_PROC = { sig:'proc' }, SIG_END = { sig:'end' };
function GotoSignal(label){ this.label = label; }

function Interp(mod, opts){
  opts = opts || {};
  this.mod = mod;
  this.wb = opts.workbook;
  this.out = [];
  this.steps = 0;
  this.maxSteps = opts.maxSteps || 4000000;
  this.inputs = (opts.inputs || []).slice();
  this.msgAnswer = opts.msgAnswer === undefined ? 6 : opts.msgAnswer;
  this.procs = {};
  this.globals = new Map();
  this.lastErr = null;
  this.appFlags = {};
  this.explicit = !!mod.explicit;
  var self = this;
  mod.procs.forEach(function(p){ self.procs[p.name.toLowerCase()] = p; });
}
var R = Interp.prototype;

R.tick = function(){ if (++this.steps > this.maxSteps) throw new VbaError(0, '実行が長すぎます。無限ループになっていませんか？'); };
R.log = function(kind, text){ this.out.push({ kind: kind, text: text }); if (this.out.length > 400) throw new VbaError(0, '出力が多すぎます（ループが止まっていない可能性があります）'); };

R.runModule = function(entry){
  var self = this;
  this.mod.consts.forEach(function(c){ self.globals.set(c.name.toLowerCase(), { v: self.eval(c.value, self.rootScope()) }); });
  this.mod.vars.forEach(function(d){ d.decls.forEach(function(x){ self.globals.set(x.name.toLowerCase(), { v: self.initialFor(x) }); }); });
  var proc = entry ? this.procs[entry.toLowerCase()] : null;
  if (!proc){
    for (var k in this.procs){ if (this.procs[k].kind === 'sub'){ proc = this.procs[k]; break; } }
  }
  if (!proc) throw new VbaError(0, '実行できる Sub が見つかりません。「Sub 名前()」〜「End Sub」で囲んでください');
  this.runProc(proc, []);
  return proc.name;
};

R.rootScope = function(){ return { vars: new Map(), withStack: [], errMode: 'off' }; };

R.initialFor = function(decl){
  var t = (decl.type || 'Variant').toLowerCase();
  if (decl.dims) return null;
  if (['long','integer','double','single','currency','byte','decimal'].indexOf(t) >= 0) return 0;
  if (t === 'string') return '';
  if (t === 'boolean') return false;
  return EMPTY;
};

R.runProc = function(proc, args){
  if ((this.depth = (this.depth||0) + 1) > 60){ this.depth--; throw new VbaError(28, '呼び出しが深すぎます（再帰が止まっていません）'); }
  var scope = { vars: new Map(), withStack: [], errMode: 'off', errLabel: null, proc: proc };
  for (var i=0;i<proc.params.length;i++){
    var p = proc.params[i];
    var v = i < args.length && args[i] !== undefined ? args[i] : (p.def ? this.eval(p.def, scope) : EMPTY);
    scope.vars.set(p.name.toLowerCase(), { v: v });
  }
  if (proc.kind === 'function') scope.vars.set(proc.name.toLowerCase(), { v: EMPTY });

  var body = proc.body, labels = {};
  for (var j=0;j<body.length;j++) if (body[j].type === 'label') labels[body[j].name.toLowerCase()] = j;

  var k = 0;
  while (k < body.length){
    try { this.execStmt(body[k], scope); k++; }
    catch (e){
      if (e === SIG_EXIT_PROC) break;
      if (e === SIG_END) { this.depth--; throw e; }
      if (e instanceof GotoSignal){
        var t = labels[String(e.label).toLowerCase()];
        if (t === undefined){ this.depth--; throw new VbaError(0, 'ラベル「' + e.label + '」が見つかりません'); }
        k = t + 1; continue;
      }
      if (e && e.isVbaError){
        this.lastErr = e;
        if (scope.errMode === 'resume'){ k++; continue; }
        if (scope.errMode === 'goto'){
          var t2 = labels[String(scope.errLabel).toLowerCase()];
          if (t2 !== undefined){ scope.errMode = 'off'; k = t2 + 1; continue; }
        }
      }
      this.depth--; throw e;
    }
  }
  this.depth--;
  if (proc.kind === 'function'){ var box = scope.vars.get(proc.name.toLowerCase()); return box ? box.v : EMPTY; }
  return EMPTY;
};

R.execBlock = function(list, scope){ for (var i=0;i<list.length;i++) this.execStmt(list[i], scope); };

R.execStmt = function(st, scope){
  this.tick();
  this.curLine = st.line || this.curLine;
  switch (st.type){
    case 'label': return;
    case 'dim': {
      for (var i=0;i<st.decls.length;i++){
        var d = st.decls[i];
        var init = d.dims ? this.makeArray(d.dims, scope, d.type) : this.initialFor(d);
        scope.vars.set(d.name.toLowerCase(), { v: init, type: d.type });
      }
      return;
    }
    case 'redim': {
      var box = this.lookupBox(st.name, scope, true);
      var arr = this.makeArray(st.dims, scope, box && box.type);
      if (st.preserve && box.v instanceof VbaArray){
        var old = box.v, n = Math.min(old.data.length, arr.data.length);
        for (var q=0;q<n;q++) arr.data[q] = old.data[q];
      }
      box.v = arr; return;
    }
    case 'const': { scope.vars.set(st.name.toLowerCase(), { v: this.eval(st.value, scope), isConst: true }); return; }
    case 'assign': return this.assign(st.target, this.eval(st.value, scope), scope);
    case 'if': {
      for (var b=0;b<st.branches.length;b++){
        if (toBool(this.eval(st.branches[b].cond, scope))) return this.execBlock(st.branches[b].body, scope);
      }
      return this.execBlock(st.elseBody || [], scope);
    }
    case 'for': {
      var from = toNum(this.eval(st.from, scope)), to = toNum(this.eval(st.to, scope));
      var step = st.step ? toNum(this.eval(st.step, scope)) : 1;
      if (step === 0) throw new VbaError(0, 'For の Step に 0 は使えません（無限ループになります）');
      var vb = this.lookupBox(st.varName, scope, true);
      for (var v = from; step > 0 ? v <= to : v >= to; v += step){
        vb.v = v; this.tick();
        try { this.execBlock(st.body, scope); }
        catch (e){ if (e === SIG_EXIT_FOR) return; throw e; }
      }
      return;
    }
    case 'foreach': {
      var items = this.enumerate(this.eval(st.coll, scope));
      var vb2 = this.lookupBox(st.varName, scope, true);
      for (var m=0;m<items.length;m++){
        vb2.v = items[m]; this.tick();
        try { this.execBlock(st.body, scope); }
        catch (e2){ if (e2 === SIG_EXIT_FOR) return; throw e2; }
      }
      return;
    }
    case 'do': {
      while (true){
        this.tick();
        if (st.pre){
          var c = toBool(this.eval(st.pre.cond, scope));
          if (st.pre.until ? c : !c) return;
        }
        try { this.execBlock(st.body, scope); }
        catch (e3){ if (e3 === SIG_EXIT_DO) return; throw e3; }
        if (st.post){
          var c2 = toBool(this.eval(st.post.cond, scope));
          if (st.post.until ? c2 : !c2) return;
        }
        if (!st.pre && !st.post) throw new VbaError(0, 'Do ... Loop に条件がありません（無限ループです）');
      }
    }
    case 'select': {
      var subj = this.eval(st.subject, scope);
      for (var s=0;s<st.cases.length;s++){
        var cs = st.cases[s];
        for (var t=0;t<cs.tests.length;t++){
          if (this.caseMatch(subj, cs.tests[t], scope)) return this.execBlock(cs.body, scope);
        }
      }
      if (st.elseBody) return this.execBlock(st.elseBody, scope);
      return;
    }
    case 'with': {
      scope.withStack.push(this.eval(st.obj, scope));
      try { this.execBlock(st.body, scope); } finally { scope.withStack.pop(); }
      return;
    }
    case 'exit': {
      if (st.what === 'for') throw SIG_EXIT_FOR;
      if (st.what === 'do') throw SIG_EXIT_DO;
      throw SIG_EXIT_PROC;
    }
    case 'endprog': throw SIG_END;
    case 'onerror': {
      scope.errMode = st.mode === 'off' ? 'off' : st.mode;
      scope.errLabel = st.label || null;
      return;
    }
    case 'resume': {
      if (st.target === '__next' || st.target === '0' || !st.target) return;
      throw new GotoSignal(st.target);
    }
    case 'goto': throw new GotoSignal(st.label);
    case 'callstmt': {
      var args = this.evalArgs(st.args, scope);
      return this.invokeCallee(st.callee, args, scope, true);
    }
    case 'exprstmt': { this.eval(st.expr, scope, true); return; }
  }
  throw new VbaError(0, '未対応の文です: ' + st.type);
};

R.caseMatch = function(subj, test, scope){
  if (test.kind === 'eq') return compare(subj, this.eval(test.value, scope)) === 0;
  if (test.kind === 'range') return compare(subj, this.eval(test.lo, scope)) >= 0 && compare(subj, this.eval(test.hi, scope)) <= 0;
  var v = this.eval(test.value, scope), c = compare(subj, v);
  switch (test.op){
    case '=': return c === 0; case '<>': return c !== 0;
    case '<': return c < 0; case '<=': return c <= 0;
    case '>': return c > 0; case '>=': return c >= 0;
  }
  return false;
};

R.makeArray = function(dims, scope, type){
  var ds = [];
  for (var i=0;i<dims.length;i++){
    var lo = toNum(this.eval(dims[i].lo, scope));
    var hi = dims[i].hi ? toNum(this.eval(dims[i].hi, scope)) : null;
    ds.push(hi === null ? { lb:0, ub:Math.floor(lo) } : { lb:Math.floor(lo), ub:Math.floor(hi) });
  }
  var total = ds.reduce(function(a,d){ return a * Math.max(0, d.ub - d.lb + 1); }, 1);
  if (total > 2000000) throw new VbaError(7, '配列が大きすぎます');
  var arr = new VbaArray(ds);
  var t = String(type || 'Variant').toLowerCase();
  if (['long','integer','double','single','currency','byte','decimal'].indexOf(t) >= 0) arr.data.fill(0);
  else if (t === 'string') arr.data.fill('');
  else if (t === 'boolean') arr.data.fill(false);
  return arr;
};

R.lookupBox = function(name, scope, create){
  var k = name.toLowerCase();
  if (scope.vars.has(k)) return scope.vars.get(k);
  if (this.globals.has(k)) return this.globals.get(k);
  if (create){
    if (this.explicit) throw new VbaError(0, '変数「' + name + '」が定義されていません。Dim ' + name + ' As ... で宣言してください');
    var b = { v: EMPTY }; scope.vars.set(k, b); return b;
  }
  return null;
};

/* --- 代入 --- */
R.assign = function(target, value, scope){
  if (target.type === 'id'){
    var box = this.lookupBox(target.name, scope, false);
    if (!box){
      if (this.explicit) throw new VbaError(0, '変数「' + target.name + '」が定義されていません。Dim ' + target.name + ' As ... で宣言してください');
      box = { v: EMPTY }; scope.vars.set(target.name.toLowerCase(), box);
    }
    if (box.isConst) throw new VbaError(0, '定数「' + target.name + '」には代入できません');
    box.v = value; return;
  }
  if (target.type === 'member'){
    var obj = target.target.type === 'withref' ? this.currentWith(scope) : this.eval(target.target, scope);
    var args = target.args ? this.evalArgs(target.args, scope) : null;
    return this.setMember(obj, target.name, args, value);
  }
  if (target.type === 'call'){
    var base = target.callee;
    if (base.type === 'id'){
      var b2 = this.lookupBox(base.name, scope, false);
      if (b2 && b2.v instanceof VbaArray){
        b2.v.set(this.evalArgs(target.args, scope).map(toNum), value); return;
      }
      if (b2 && b2.v && b2.v._t === 'Dict'){ dictSet(b2.v, this.eval(target.args[0], scope), value); return; }
    }
    var o = this.eval(target.callee, scope);
    var a = this.evalArgs(target.args, scope);
    if (o && o._t === 'Dict'){ dictSet(o, a[0], value); return; }
    var resolved = this.callObject(o, a);
    if (resolved && resolved._t === 'Range'){ setRangeValue(resolved, value); return; }
    throw new VbaError(0, 'この式には代入できません');
  }
  throw new VbaError(0, 'この式には代入できません');
};

R.currentWith = function(scope){
  if (!scope.withStack.length) throw new VbaError(0, 'With ブロックの外で「.」から始まる式が使われています');
  return scope.withStack[scope.withStack.length-1];
};

R.evalArgs = function(args, scope){
  var out = [], self = this;
  (args||[]).forEach(function(a){
    if (a === null){ out.push(undefined); return; }
    if (a && a.type === 'named'){ var v = self.eval(a.value, scope); v = v; out.push(v); out.named = out.named || {}; out.named[a.name.toLowerCase()] = v; return; }
    out.push(self.eval(a, scope));
  });
  return out;
};

/* --- 式評価 --- */
R.eval = function(node, scope, asStatement){
  this.tick();
  switch (node.type){
    case 'lit': return node.value;
    case 'id': {
      var box = this.lookupBox(node.name, scope, false);
      if (box) return box.v;
      var k = node.name.toLowerCase();
      if (CONSTS.hasOwnProperty(k)) return CONSTS[k];
      var g = this.globalObject(k);
      if (g !== undefined) return g;
      if (this.procs[k]) return this.runProc(this.procs[k], []);
      if (BUILTINS[k]) return BUILTINS[k].call(this, []);
      throw new VbaError(0, '「' + node.name + '」が定義されていません（Dim を書き忘れていませんか）');
    }
    case 'withref': return this.currentWith(scope);
    case 'new': return this.newObject(node.className);
    case 'un': {
      if (node.op === '-') return -toNum(this.eval(node.e, scope));
      var v = this.eval(node.e, scope);
      if (typeof v === 'boolean') return !v;
      return ~toNum(v);
    }
    case 'bin': return this.binop(node, scope);
    case 'member': {
      var obj = node.target.type === 'withref' ? this.currentWith(scope) : this.eval(node.target, scope);
      var args = node.args ? this.evalArgs(node.args, scope) : null;
      return this.getMember(obj, node.name, args);
    }
    case 'call': {
      var a = this.evalArgs(node.args, scope);
      return this.invokeCallee(node.callee, a, scope, asStatement);
    }
  }
  throw new VbaError(0, '未対応の式です');
};

R.invokeCallee = function(callee, args, scope, asStatement){
  if (callee.type === 'id'){
    var k = callee.name.toLowerCase();
    var box = this.lookupBox(callee.name, scope, false);
    if (box && box.v instanceof VbaArray) return box.v.get(args.map(toNum));
    if (box && box.v && box.v._t) return this.callObject(box.v, args);
    if (this.procs[k]) return this.runProc(this.procs[k], args);
    if (BUILTINS[k]) return BUILTINS[k].call(this, args, asStatement);
    var g = this.globalObject(k, args);
    if (g !== undefined) return g;
    throw new VbaError(0, '「' + callee.name + '」という関数は見つかりません');
  }
  if (callee.type === 'member'){
    var obj = callee.target.type === 'withref' ? this.currentWith(scope) : this.eval(callee.target, scope);
    return this.getMember(obj, callee.name, args);
  }
  var o = this.eval(callee, scope);
  return this.callObject(o, args);
};

R.callObject = function(o, args){
  if (o && o._t === 'Range'){
    if (!args.length) return o;
    return this.getMember(o, 'Item', args);
  }
  if (o && o._t === 'Sheets') return this.getMember(o, 'Item', args);
  if (o && o._t === 'Dict') return dictGet(o, args[0]);
  if (o && o._t === 'Coll') return collItem(o, args[0]);
  if (o instanceof VbaArray) return o.get(args.map(toNum));
  if (!args.length) return o;
  throw new VbaError(0, 'この値は呼び出せません');
};

function compare(a, b){
  if (a && a._t === 'Range') a = rangeValue(a);
  if (b && b._t === 'Range') b = rangeValue(b);
  if (isDate(a) || isDate(b)) return Math.sign(toNum(a) - toNum(b));
  if (typeof a === 'string' || typeof b === 'string'){
    if (typeof a === 'number' || typeof b === 'number'){
      var na = Number(String(a).replace(/,/g,'')), nb = Number(String(b).replace(/,/g,''));
      if (!isNaN(na) && !isNaN(nb)) return Math.sign(na - nb);
    }
    var sa = toStr(a), sb = toStr(b);
    return sa < sb ? -1 : sa > sb ? 1 : 0;
  }
  return Math.sign(toNum(a) - toNum(b));
}

R.binop = function(node, scope){
  var op = node.op;
  if (op === 'and'){ var l = this.eval(node.l, scope); if (typeof l === 'boolean' && !l) return false; var r = this.eval(node.r, scope); return (typeof l === 'boolean' || typeof r === 'boolean') ? (toBool(l) && toBool(r)) : (toNum(l) & toNum(r)); }
  if (op === 'or'){ var l2 = this.eval(node.l, scope); if (typeof l2 === 'boolean' && l2) return true; var r2 = this.eval(node.r, scope); return (typeof l2 === 'boolean' || typeof r2 === 'boolean') ? (toBool(l2) || toBool(r2)) : (toNum(l2) | toNum(r2)); }
  var a = this.eval(node.l, scope), b = this.eval(node.r, scope);
  if (a && a._t === 'Range' && op !== 'is') a = rangeValue(a);
  if (b && b._t === 'Range' && op !== 'is') b = rangeValue(b);
  switch (op){
    case '&': return toStr(a) + toStr(b);
    case '+':
      if (typeof a === 'string' && typeof b === 'string') return a + b;
      if (isDate(a) && typeof b === 'number') return serialToDate(dateToSerial(a) + b);
      return toNum(a) + toNum(b);
    case '-':
      if (isDate(a) && isDate(b)) return dateToSerial(a) - dateToSerial(b);
      if (isDate(a)) return serialToDate(dateToSerial(a) - toNum(b));
      return toNum(a) - toNum(b);
    case '*': return toNum(a) * toNum(b);
    case '/': { var d = toNum(b); if (d === 0) throw new VbaError(11, '0 で割ることはできません'); return toNum(a) / d; }
    case '\\': { var d2 = toNum(b); if (d2 === 0) throw new VbaError(11, '0 で割ることはできません'); return Math.trunc(Math.trunc(toNum(a)) / Math.trunc(d2)); }
    case 'mod': { var d3 = toNum(b); if (d3 === 0) throw new VbaError(11, '0 で割ることはできません'); return Math.trunc(toNum(a)) % Math.trunc(d3); }
    case '^': return Math.pow(toNum(a), toNum(b));
    case '=': return compare(a,b) === 0;
    case '<>': return compare(a,b) !== 0;
    case '<': return compare(a,b) < 0;
    case '<=': return compare(a,b) <= 0;
    case '>': return compare(a,b) > 0;
    case '>=': return compare(a,b) >= 0;
    case 'is': return a === b || (a === NOTHING && b === NOTHING);
    case 'like': {
      var pat = toStr(b).replace(/[.+^${}()|[\]\\]/g,'\\$&').replace(/\*/g,'.*').replace(/\?/g,'.').replace(/#/g,'[0-9]');
      return new RegExp('^' + pat + '$', 'i').test(toStr(a));
    }
  }
  throw new VbaError(0, '未対応の演算子: ' + op);
};

R.enumerate = function(v){
  if (v instanceof VbaArray) return v.data.slice();
  if (v && v._t === 'Range') return rangeCells(v);
  if (v && v._t === 'Coll') return v.items.slice();
  if (v && v._t === 'Sheets') return this.wb.sheets.map(mkSheetObj);
  if (v && v._t === 'Dict') return v.order.slice();
  if (v && v._t === 'StrColl') return v.items.slice();
  throw new VbaError(0, 'For Each で回せない値です');
};

/* ---------- Dictionary / Collection ---------- */
function dictKey(k){ return typeof k === 'string' ? 's' + k.toLowerCase() : 'n' + toStr(k); }
function dictGet(d, k){ var e = d.map.get(dictKey(k)); return e ? e.v : EMPTY; }
function dictSet(d, k, v){
  var kk = dictKey(k);
  if (!d.map.has(kk)) d.order.push(k);
  d.map.set(kk, { k: k, v: v });
}
function collItem(c, i){
  if (typeof i === 'string'){ var e = c.keys.get(i.toLowerCase()); if (!e) throw new VbaError(5, 'キー「' + i + '」は Collection にありません'); return e; }
  var n = Math.floor(toNum(i));
  if (n < 1 || n > c.items.length) throw new VbaError(9, 'インデックスが有効範囲にありません（' + n + '）');
  return c.items[n-1];
}

R.newObject = function(cls){
  var c = String(cls).toLowerCase();
  if (c === 'collection') return { _t:'Coll', items: [], keys: new Map() };
  if (c === 'dictionary' || c === 'scripting.dictionary') return { _t:'Dict', map: new Map(), order: [] };
  throw new VbaError(429, '「' + cls + '」は作れません');
};

/* ---------- グローバル ---------- */
R.globalObject = function(k, args){
  switch (k){
    case 'application': return { _t:'App' };
    case 'thisworkbook': case 'activeworkbook': return { _t:'Book' };
    case 'activesheet': return mkSheetObj(this.wb.active);
    case 'worksheets': case 'sheets':
      if (args && args.length) return this.getMember({ _t:'Sheets' }, 'Item', args);
      return { _t:'Sheets' };
    case 'range': if (args) return parseAddressArgs(this.wb.active, args); return undefined;
    case 'cells': return this.getMember(mkSheetObj(this.wb.active), 'Cells', args || []);
    case 'rows': return { _t:'Rows', sheet: this.wb.active };
    case 'columns': return { _t:'Cols', sheet: this.wb.active };
    case 'worksheetfunction': return { _t:'WF' };
    case 'err': return { _t:'Err' };
    case 'debug': return { _t:'Debug' };
    case 'selection': return mkRange(this.wb.active, 1,1,1,1);
  }
  return undefined;
};

function parseAddressArgs(sheet, args){
  if (args.length >= 2 && args[0] && args[0]._t === 'Range' && args[1] && args[1]._t === 'Range'){
    var a = args[0], b = args[1];
    return mkRange(sheet, Math.min(a.r1,b.r1), Math.min(a.c1,b.c1), Math.max(a.r2,b.r2), Math.max(a.c2,b.c2));
  }
  if (args[0] && args[0]._t === 'Range') return args[0];
  return parseAddress(sheet, toStr(args[0]));
}
function colArg(v){
  if (typeof v === 'string' && /^[A-Za-z]{1,3}$/.test(v.trim())) return letterCol(v.trim());
  return Math.floor(toNum(v));
}

/* ---------- メンバー取得 ---------- */
R.getMember = function(obj, rawName, args){
  var name = String(rawName).toLowerCase();
  if (obj === undefined || obj === null || obj === EMPTY || obj === NOTHING)
    throw new VbaError(91, 'オブジェクトが設定されていません（Set を忘れていませんか）「.' + rawName + '」');
  var t = obj._t;

  if (t === 'App'){
    switch (name){
      case 'worksheetfunction': return { _t:'WF' };
      case 'username': return '学習ユーザー';
      case 'pathseparator': return '/';
      case 'screenupdating': case 'displayalerts': case 'enableevents': case 'calculation':
      case 'cutcopymode': case 'statusbar': case 'displaystatusbar':
        return this.appFlags[name] === undefined ? true : this.appFlags[name];
      case 'workbooks': return { _t:'Books' };
      case 'activesheet': return mkSheetObj(this.wb.active);
      case 'inputbox': return BUILTINS['inputbox'].call(this, args||[]);
      case 'wait': case 'calculate': case 'ontime': return EMPTY;
      case 'version': return '16.0';
      default:
        if (WF[name]) return WF[name].call(this, args||[], true);
        return EMPTY;
    }
  }
  if (t === 'WF'){
    if (WF[name]) return WF[name].call(this, args||[], false);
    throw new VbaError(1004, 'WorksheetFunction.' + rawName + ' は未対応です');
  }
  if (t === 'Err'){
    switch (name){
      case 'number': return this.lastErr ? this.lastErr.number : 0;
      case 'description': return this.lastErr ? this.lastErr.description : '';
      case 'source': return 'VBA練習';
      case 'clear': this.lastErr = null; return EMPTY;
      case 'raise': { var n = args && args.length ? toNum(args[0]) : 5; throw new VbaError(n, args && args.length>2 ? toStr(args[2]) : 'Err.Raise で発生させたエラー'); }
    }
    return EMPTY;
  }
  if (t === 'Debug'){
    if (name === 'print'){ this.log('debug', (args||[]).map(function(a){ return a === undefined ? '' : toStr(a); }).join(' ')); return EMPTY; }
    return EMPTY;
  }
  if (t === 'Book'){
    switch (name){
      case 'worksheets': case 'sheets':
        if (args && args.length) return this.getMember({ _t:'Sheets' }, 'Item', args);
        return { _t:'Sheets' };
      case 'name': return this.wb.name;
      case 'path': return '/Users/you/Desktop';
      case 'fullname': return '/Users/you/Desktop/' + this.wb.name;
      case 'activesheet': return mkSheetObj(this.wb.active);
      case 'save': case 'saveas': case 'close': case 'activate': return EMPTY;
    }
    return EMPTY;
  }
  if (t === 'Sheets'){
    switch (name){
      case 'count': return this.wb.sheets.length;
      case 'item': {
        var a0 = args && args.length ? args[0] : 1;
        var s = typeof a0 === 'string' ? this.wb.byName(a0) : this.wb.sheets[Math.floor(toNum(a0))-1];
        if (!s) throw new VbaError(9, 'シート「' + toStr(a0) + '」が見つかりません');
        return mkSheetObj(s);
      }
      case 'add': {
        var ns = new Sheet('Sheet' + (this.wb.sheets.length + 1));
        this.wb.sheets.push(ns); this.wb.active = ns;
        return mkSheetObj(ns);
      }
    }
    return EMPTY;
  }
  if (t === 'Sheet'){
    var sh = obj.sheet;
    switch (name){
      case 'name': return sh.name;
      case 'range': return parseAddressArgs(sh, args || []);
      case 'cells': {
        if (!args || !args.length) return mkRange(sh, 1, 1, MAXROW, MAXCOL);
        if (args.length === 1){ var idx = Math.floor(toNum(args[0])); var c = ((idx-1) % MAXCOL) + 1, r = Math.floor((idx-1)/MAXCOL)+1; return mkRange(sh,r,c,r,c); }
        var rr = Math.floor(toNum(args[0])), cc = colArg(args[1]);
        if (rr < 1 || cc < 1) throw new VbaError(1004, 'Cells の行・列は 1 以上です（' + rr + ', ' + cc + '）');
        return mkRange(sh, rr, cc, rr, cc);
      }
      case 'rows': if (args && args.length){ var rn = Math.floor(toNum(args[0])); return mkRange(sh, rn, 1, rn, MAXCOL); } return { _t:'Rows', sheet: sh };
      case 'columns': if (args && args.length){ var cn = colArg(args[0]); return mkRange(sh, 1, cn, MAXROW, cn); } return { _t:'Cols', sheet: sh };
      case 'usedrange': { var b = sh.bounds(); return mkRange(sh, 1, 1, Math.max(1,b.rows), Math.max(1,b.cols)); }
      case 'activate': case 'select': this.wb.active = sh; return EMPTY;
      case 'delete': {
        var i = this.wb.sheets.indexOf(sh);
        if (i >= 0 && this.wb.sheets.length > 1){ this.wb.sheets.splice(i,1); if (this.wb.active === sh) this.wb.active = this.wb.sheets[0]; }
        return EMPTY;
      }
      case 'tab': return { _t:'Stub' };
      case 'autofiltermode': return false;
      case 'pagesetup': return { _t:'Stub' };
      case 'index': return this.wb.sheets.indexOf(sh) + 1;
      case 'visible': return true;
    }
    return { _t:'Stub' };
  }
  if (t === 'Rows'){ if (name === 'count') return MAXROW; if (name === 'autofit' || name === 'hidden') return EMPTY; return { _t:'Stub' }; }
  if (t === 'Cols'){ if (name === 'count') return MAXCOL; if (name === 'autofit' || name === 'hidden') return EMPTY; return { _t:'Stub' }; }
  if (t === 'RangeRows'){ if (name === 'count') return obj.n; if (name === 'item'){ var k = Math.floor(toNum(args[0])); var g = obj.range; return mkRange(g.sheet, g.r1+k-1, g.c1, g.r1+k-1, g.c2); } return { _t:'Stub' }; }
  if (t === 'RangeCols'){ if (name === 'count') return obj.n; if (name === 'item'){ var k2 = Math.floor(toNum(args[0])); var g2 = obj.range; return mkRange(g2.sheet, g2.r1, g2.c1+k2-1, g2.r2, g2.c1+k2-1); } return { _t:'Stub' }; }

  if (t === 'Range'){
    var rg = obj;
    switch (name){
      case 'value': case 'value2': case 'formula': case 'formular1c1': return rangeValue(rg);
      case 'text': { var cl = rg.sheet.cell(rg.r1,rg.c1,false); return applyFormat(rangeValue(rg), cl && cl.fmt); }
      case 'row': return rg.r1;
      case 'column': return rg.c1;
      case 'count': return (rg.r2-rg.r1+1) * (rg.c2-rg.c1+1);
      case 'address': return '$' + colLetter(rg.c1) + '$' + rg.r1 + (rg.r1===rg.r2 && rg.c1===rg.c2 ? '' : ':$' + colLetter(rg.c2) + '$' + rg.r2);
      case 'rows': return { _t:'RangeRows', range: rg, n: rg.r2-rg.r1+1 };
      case 'columns': return { _t:'RangeCols', range: rg, n: rg.c2-rg.c1+1 };
      case 'font': return { _t:'Font', range: rg };
      case 'interior': return { _t:'Interior', range: rg };
      case 'borders': return { _t:'Stub' };
      case 'numberformat': case 'numberformatlocal': { var c0 = rg.sheet.cell(rg.r1,rg.c1,false); return (c0 && c0.fmt) || 'General'; }
      case 'horizontalalignment': case 'verticalalignment': return -4108;
      case 'offset': {
        var dr = args && args[0] !== undefined ? Math.floor(toNum(args[0])) : 0;
        var dc = args && args[1] !== undefined ? Math.floor(toNum(args[1])) : 0;
        if (rg.r1+dr < 1 || rg.c1+dc < 1) throw new VbaError(1004, 'Offset の結果がシートの外に出ました');
        return mkRange(rg.sheet, rg.r1+dr, rg.c1+dc, rg.r2+dr, rg.c2+dc);
      }
      case 'resize': {
        var nr = args && args[0] !== undefined ? Math.floor(toNum(args[0])) : (rg.r2-rg.r1+1);
        var nc = args && args[1] !== undefined ? Math.floor(toNum(args[1])) : (rg.c2-rg.c1+1);
        if (nr < 1 || nc < 1) throw new VbaError(1004, 'Resize の行数・列数は 1 以上です');
        return mkRange(rg.sheet, rg.r1, rg.c1, rg.r1+nr-1, rg.c1+nc-1);
      }
      case 'end': {
        var d = args && args.length ? args[0] : 'up';
        return endMove(rg, typeof d === 'string' ? d : 'up');
      }
      case 'currentregion': return currentRegion(rg);
      case 'entirerow': return mkRange(rg.sheet, rg.r1, 1, rg.r2, MAXCOL);
      case 'entirecolumn': return mkRange(rg.sheet, 1, rg.c1, MAXROW, rg.c2);
      case 'item': {
        if (args.length >= 2) return mkRange(rg.sheet, rg.r1+Math.floor(toNum(args[0]))-1, rg.c1+colArg(args[1])-1, rg.r1+Math.floor(toNum(args[0]))-1, rg.c1+colArg(args[1])-1);
        var w = rg.c2-rg.c1+1, i2 = Math.floor(toNum(args[0]))-1;
        return mkRange(rg.sheet, rg.r1+Math.floor(i2/w), rg.c1+(i2%w), rg.r1+Math.floor(i2/w), rg.c1+(i2%w));
      }
      case 'clearcontents': { forEachCell(rg, function(sh2,r,c){ sh2.setV(r,c,EMPTY); }); return EMPTY; }
      case 'clear': { forEachCell(rg, function(sh3,r,c){ var x = sh3.cell(r,c,true); x.v = EMPTY; delete x.bold; delete x.fill; delete x.fcolor; delete x.fmt; delete x.italic; }); return EMPTY; }
      case 'clearformats': { forEachCell(rg, function(sh4,r,c){ var x = sh4.cell(r,c,true); delete x.bold; delete x.fill; delete x.fcolor; delete x.fmt; }); return EMPTY; }
      case 'select': case 'activate': case 'autofit': case 'copy': case 'merge': case 'autofilter': return EMPTY;
      case 'worksheet': case 'parent': return mkSheetObj(rg.sheet);
      case 'find': return NOTHING;
    }
    return { _t:'Stub' };
  }
  if (t === 'Font' || t === 'Interior'){
    var cl2 = obj.range.sheet.cell(obj.range.r1, obj.range.c1, false);
    if (t === 'Font'){
      switch (name){
        case 'bold': return !!(cl2 && cl2.bold);
        case 'italic': return !!(cl2 && cl2.italic);
        case 'size': return (cl2 && cl2.size) || 11;
        case 'color': return (cl2 && cl2.fcolor !== undefined) ? cl2.fcolor : 0;
        case 'name': return 'Meiryo UI';
      }
    } else {
      switch (name){
        case 'color': return (cl2 && cl2.fill !== undefined) ? cl2.fill : 16777215;
        case 'colorindex': return -4142;
      }
    }
    return EMPTY;
  }
  if (t === 'Dict'){
    switch (name){
      case 'exists': return obj.map.has(dictKey(args[0]));
      case 'add': { if (obj.map.has(dictKey(args[0]))) throw new VbaError(457, 'このキーは既に登録されています'); dictSet(obj, args[0], args[1]); return EMPTY; }
      case 'item': { if (args.length > 1){ dictSet(obj, args[0], args[1]); return EMPTY; } return dictGet(obj, args[0]); }
      case 'remove': { obj.map.delete(dictKey(args[0])); obj.order = obj.order.filter(function(k){ return dictKey(k) !== dictKey(args[0]); }); return EMPTY; }
      case 'removeall': obj.map.clear(); obj.order = []; return EMPTY;
      case 'count': return obj.order.length;
      case 'keys': return new VbaArray([{lb:0,ub:Math.max(0,obj.order.length-1)}], obj.order.slice());
      case 'items': { var vs = obj.order.map(function(k){ return dictGet(obj,k); }); return new VbaArray([{lb:0,ub:Math.max(0,vs.length-1)}], vs); }
      case 'comparemode': return 0;
    }
    return EMPTY;
  }
  if (t === 'Coll'){
    switch (name){
      case 'add': {
        var item = args[0], key = args[1];
        if (key !== undefined && key !== EMPTY){
          var kk = toStr(key).toLowerCase();
          if (obj.keys.has(kk)) throw new VbaError(457, 'このキーは既に登録されています');
          obj.keys.set(kk, item);
        }
        obj.items.push(item); return EMPTY;
      }
      case 'count': return obj.items.length;
      case 'item': return collItem(obj, args[0]);
      case 'remove': { var i3 = Math.floor(toNum(args[0])); obj.items.splice(i3-1,1); return EMPTY; }
    }
    return EMPTY;
  }
  if (t === 'Stub') return { _t:'Stub' };
  if (obj instanceof VbaArray){
    if (name === 'count' || name === 'length') return obj.data.length;
  }
  throw new VbaError(438, 'このオブジェクトに「' + rawName + '」はありません');
};

function forEachCell(rg, fn){
  ACCESS.n++;
  var r2 = Math.min(rg.r2, Math.max(rg.sheet.bounds().rows, rg.r1 + 2000));
  var c2 = Math.min(rg.c2, Math.max(rg.sheet.bounds().cols, rg.c1 + 100));
  for (var r=rg.r1;r<=r2;r++) for (var c=rg.c1;c<=c2;c++) fn(rg.sheet, r, c);
}

R.setMember = function(obj, rawName, args, value){
  var name = String(rawName).toLowerCase();
  if (obj === undefined || obj === null || obj === EMPTY || obj === NOTHING)
    throw new VbaError(91, 'オブジェクトが設定されていません「.' + rawName + '」');
  var t = obj._t;
  if (t === 'App'){ this.appFlags[name] = value; return; }
  if (t === 'Sheet'){
    if (name === 'name'){
      var nm = toStr(value);
      if (this.wb.byName(nm) && this.wb.byName(nm) !== obj.sheet) throw new VbaError(1004, '同じ名前のシートが既にあります：' + nm);
      obj.sheet.name = nm; return;
    }
    if (name === 'range'){ setRangeValue(parseAddressArgs(obj.sheet, args||[]), value); return; }
    if (name === 'cells'){ setRangeValue(this.getMember(obj, 'Cells', args||[]), value); return; }
    return;
  }
  if (t === 'Range'){
    var rg = obj;
    switch (name){
      case 'value': case 'value2': case 'formula': setRangeValue(rg, value); return;
      case 'numberformat': case 'numberformatlocal': forEachCell(rg, function(sh,r,c){ sh.cell(r,c,true).fmt = toStr(value); }); return;
      case 'offset': case 'resize': case 'item': { setRangeValue(this.getMember(rg, rawName, args||[]), value); return; }
      case 'horizontalalignment': forEachCell(rg, function(sh,r,c){ sh.cell(r,c,true).halign = toNum(value); }); return;
      case 'columnwidth': case 'rowheight': return;
      default: return;
    }
  }
  if (t === 'Font'){
    var rg2 = obj.range;
    forEachCell(rg2, function(sh,r,c){
      var x = sh.cell(r,c,true);
      if (name === 'bold') x.bold = toBool(value);
      else if (name === 'italic') x.italic = toBool(value);
      else if (name === 'size') x.size = toNum(value);
      else if (name === 'color') x.fcolor = toNum(value);
      else if (name === 'colorindex') x.fcolor = 0;
    });
    return;
  }
  if (t === 'Interior'){
    var rg3 = obj.range;
    forEachCell(rg3, function(sh,r,c){
      var x = sh.cell(r,c,true);
      if (name === 'color') x.fill = toNum(value);
      else if (name === 'colorindex') x.fill = 16777215;
    });
    return;
  }
  if (t === 'Dict'){ if (name === 'item'){ dictSet(obj, args[0], value); return; } return; }
  if (t === 'Stub') return;
  return;
};

/* ---------- WorksheetFunction ---------- */
function numsOf(args){
  var out = [];
  args.forEach(function pick(a){
    if (a === undefined) return;
    if (a && a._t === 'Range'){ rangeFlatValues(a).forEach(function(v){ if (typeof v === 'number') out.push(v); else if (typeof v === 'string' && v.trim() !== '' && !isNaN(Number(v.replace(/,/g,'')))) out.push(Number(v.replace(/,/g,''))); else if (isDate(v)) out.push(dateToSerial(v)); }); return; }
    if (a instanceof VbaArray){ a.data.forEach(function(v){ if (typeof v === 'number') out.push(v); }); return; }
    if (a === EMPTY) return;
    out.push(toNum(a));
  });
  return out;
}
var WF = {
  sum: function(a){ return numsOf(a).reduce(function(x,y){ return x+y; }, 0); },
  count: function(a){ return numsOf(a).length; },
  counta: function(a){ var n = 0; a.forEach(function(x){ if (x && x._t === 'Range') rangeFlatValues(x).forEach(function(v){ if (!(v===EMPTY||v===''||v===undefined)) n++; }); else if (!(x===EMPTY||x==='')) n++; }); return n; },
  max: function(a){ var n = numsOf(a); if (!n.length) return 0; return Math.max.apply(null,n); },
  min: function(a){ var n = numsOf(a); if (!n.length) return 0; return Math.min.apply(null,n); },
  average: function(a){ var n = numsOf(a); if (!n.length) throw new VbaError(1004, 'Average: 数値がありません'); return n.reduce(function(x,y){return x+y;},0)/n.length; },
  round: function(a){ var d = a[1] === undefined ? 0 : toNum(a[1]); var m = Math.pow(10,d); return Math.round(toNum(a[0])*m)/m; },
  large: function(a){ var n = numsOf([a[0]]).sort(function(x,y){return y-x;}); return n[Math.floor(toNum(a[1]))-1]; },
  small: function(a){ var n = numsOf([a[0]]).sort(function(x,y){return x-y;}); return n[Math.floor(toNum(a[1]))-1]; },
  countif: function(a){
    var rg = a[0], crit = a[1], n = 0;
    rangeFlatValues(rg).forEach(function(v){ if (matchCrit(v, crit)) n++; });
    return n;
  },
  sumif: function(a){
    var rg = a[0], crit = a[1], sumRg = a[2] || a[0], total = 0;
    var vals = rangeFlatValues(rg), svals = rangeFlatValues(sumRg);
    for (var i=0;i<vals.length;i++) if (matchCrit(vals[i], crit)) total += (typeof svals[i] === 'number' ? svals[i] : 0);
    return total;
  },
  vlookup: function(a, soft){
    var key = a[0], rg = a[1], col = Math.floor(toNum(a[2]));
    var exact = a[3] === undefined ? true : !toBool(a[3]);
    for (var r=rg.r1;r<=Math.min(rg.r2, 200000);r++){
      var v = rg.sheet.getV(r, rg.c1);
      if (compare(v, key) === 0) return rg.sheet.getV(r, rg.c1 + col - 1);
    }
    if (soft) return new ErrValue(2042);
    throw new VbaError(1004, 'VLookup: 「' + toStr(key) + '」が見つかりません');
  },
  match: function(a, soft){
    var key = a[0], rg = a[1];
    var i = 1;
    for (var r=rg.r1;r<=Math.min(rg.r2,200000);r++){ if (compare(rg.sheet.getV(r,rg.c1), key) === 0) return i; i++; }
    if (soft) return new ErrValue(2042);
    throw new VbaError(1004, 'Match: 見つかりません');
  },
  trim: function(a){ return toStr(a[0]).trim().replace(/\s+/g,' '); },
  proper: function(a){ return toStr(a[0]).replace(/\w\S*/g, function(s){ return s[0].toUpperCase() + s.slice(1).toLowerCase(); }); }
};
function matchCrit(v, crit){
  if (typeof crit === 'string'){
    var m = /^(>=|<=|<>|>|<|=)(.*)$/.exec(crit.trim());
    if (m){
      var other = isNaN(Number(m[2])) ? m[2] : Number(m[2]);
      var c = compare(v, other);
      switch (m[1]){ case '>=': return c>=0; case '<=': return c<=0; case '<>': return c!==0; case '>': return c>0; case '<': return c<0; default: return c===0; }
    }
  }
  return compare(v, crit) === 0;
}

/* ---------- 組み込み関数 ---------- */
function arg(a,i,d){ return (a && a[i] !== undefined && a[i] !== null) ? a[i] : d; }
var BUILTINS = {
  msgbox: function(a){
    var body = a[0] === undefined ? '' : toStr(a[0]);
    var title = a[2] !== undefined ? toStr(a[2]) : null;
    this.log('msgbox', (title ? '［' + title + '］\n' : '') + body);
    return this.msgAnswer;
  },
  inputbox: function(a){
    var prompt = a[0] === undefined ? '' : toStr(a[0]);
    var v = this.inputs.length ? this.inputs.shift() : '';
    this.log('input', prompt + '  →  「' + v + '」');
    return v;
  },
  createobject: function(a){ return this.newObject(toStr(a[0])); },
  format: function(a){ return applyFormat(a[0], a[1] === undefined ? 'General' : toStr(a[1])); },
  len: function(a){ return toStr(a[0]).length; },
  left: function(a){ return toStr(a[0]).slice(0, Math.max(0,Math.floor(toNum(a[1])))); },
  right: function(a){ var n = Math.max(0,Math.floor(toNum(a[1]))); var s = toStr(a[0]); return n === 0 ? '' : s.slice(-n); },
  mid: function(a){ var s = toStr(a[0]), st = Math.floor(toNum(a[1])); var ln = a[2] === undefined ? undefined : Math.floor(toNum(a[2])); return ln === undefined ? s.slice(st-1) : s.substr(st-1, ln); },
  trim: function(a){ return toStr(a[0]).replace(/^[\s　]+|[\s　]+$/g,''); },
  ltrim: function(a){ return toStr(a[0]).replace(/^[\s　]+/,''); },
  rtrim: function(a){ return toStr(a[0]).replace(/[\s　]+$/,''); },
  ucase: function(a){ return toStr(a[0]).toUpperCase(); },
  lcase: function(a){ return toStr(a[0]).toLowerCase(); },
  instr: function(a){
    var start = 1, s, sub;
    if (a.length >= 3 && typeof a[0] === 'number'){ start = Math.floor(toNum(a[0])); s = toStr(a[1]); sub = toStr(a[2]); }
    else { s = toStr(a[0]); sub = toStr(a[1]); }
    if (sub === '') return start;
    return s.indexOf(sub, start-1) + 1;
  },
  instrrev: function(a){ return toStr(a[0]).lastIndexOf(toStr(a[1])) + 1; },
  replace: function(a){ var s = toStr(a[0]), f = toStr(a[1]), r = toStr(a[2]); if (f === '') return s; return s.split(f).join(r); },
  split: function(a){
    var s = toStr(a[0]), d = a[1] === undefined ? ' ' : toStr(a[1]);
    var parts = s.split(d);
    return new VbaArray([{lb:0,ub:parts.length-1}], parts);
  },
  join: function(a){
    var arr = a[0], d = a[1] === undefined ? ' ' : toStr(a[1]);
    if (!(arr instanceof VbaArray)) throw new VbaError(13, 'Join には配列が必要です');
    return arr.data.map(toStr).join(d);
  },
  array: function(a){ return new VbaArray([{lb:0,ub:Math.max(0,a.length-1)}], a.slice()); },
  ubound: function(a){
    var arr = a[0], d = a[1] === undefined ? 1 : Math.floor(toNum(a[1]));
    if (!(arr instanceof VbaArray)) throw new VbaError(13, 'UBound には配列が必要です');
    if (d > arr.dims.length) throw new VbaError(9, 'その次元は配列にありません');
    return arr.dims[d-1].ub;
  },
  lbound: function(a){
    var arr = a[0], d = a[1] === undefined ? 1 : Math.floor(toNum(a[1]));
    if (!(arr instanceof VbaArray)) throw new VbaError(13, 'LBound には配列が必要です');
    return arr.dims[d-1].lb;
  },
  isnumeric: function(a){
    var v = a[0];
    if (typeof v === 'number') return true;
    if (typeof v === 'boolean') return true;
    if (v === EMPTY || v === undefined) return false;
    if (isDate(v)) return false;
    if (v && v._t === 'Range') return BUILTINS.isnumeric.call(this, [rangeValue(v)]);
    var s = toStr(v).trim().replace(/,/g,'');
    return s !== '' && !isNaN(Number(s));
  },
  isempty: function(a){ var v = a[0]; if (v && v._t === 'Range') v = rangeValue(v); return v === EMPTY || v === undefined || v === null; },
  isdate: function(a){ var v = a[0]; if (v && v._t === 'Range') v = rangeValue(v); if (isDate(v)) return true; if (typeof v !== 'string') return false; return !isNaN(Date.parse(v.replace(/-/g,'/'))); },
  isarray: function(a){ return a[0] instanceof VbaArray; },
  iserror: function(a){ return !!(a[0] && a[0]._t === 'ErrorValue'); },
  isnull: function(a){ return a[0] === null; },
  isobject: function(a){ return !!(a[0] && a[0]._t); },
  cstr: function(a){ return toStr(a[0]); },
  clng: function(a){ var n = toNum(a[0]); return Math.round(n); },
  cint: function(a){ return Math.round(toNum(a[0])); },
  cdbl: function(a){ return toNum(a[0]); },
  csng: function(a){ return toNum(a[0]); },
  ccur: function(a){ return toNum(a[0]); },
  cbool: function(a){ return toBool(a[0]); },
  cdate: function(a){
    var v = a[0];
    if (isDate(v)) return v;
    if (typeof v === 'number') return serialToDate(v);
    var s = toStr(v).replace(/-/g,'/');
    var d = new Date(s);
    if (isNaN(d.getTime())) throw new VbaError(13, '「' + toStr(v) + '」は日付になりません');
    return d;
  },
  val: function(a){ var m = /^[\s]*[-+]?[0-9]*\.?[0-9]+/.exec(toStr(a[0])); return m ? parseFloat(m[0]) : 0; },
  int: function(a){ return Math.floor(toNum(a[0])); },
  fix: function(a){ return Math.trunc(toNum(a[0])); },
  abs: function(a){ return Math.abs(toNum(a[0])); },
  sgn: function(a){ return Math.sign(toNum(a[0])); },
  sqr: function(a){ return Math.sqrt(toNum(a[0])); },
  round: function(a){ var d = a[1] === undefined ? 0 : Math.floor(toNum(a[1])); var m = Math.pow(10,d); return Math.round(toNum(a[0])*m)/m; },
  rnd: function(){ return Math.random(); },
  timer: function(){ return (Date.now() % 86400000) / 1000; },
  now: function(){ return new Date(); },
  date: function(){ var d = new Date(); return new Date(d.getFullYear(), d.getMonth(), d.getDate()); },
  time: function(){ return new Date(); },
  year: function(a){ return BUILTINS.cdate.call(this,a).getFullYear(); },
  month: function(a){ return BUILTINS.cdate.call(this,a).getMonth()+1; },
  day: function(a){ return BUILTINS.cdate.call(this,a).getDate(); },
  hour: function(a){ return BUILTINS.cdate.call(this,a).getHours(); },
  minute: function(a){ return BUILTINS.cdate.call(this,a).getMinutes(); },
  second: function(a){ return BUILTINS.cdate.call(this,a).getSeconds(); },
  weekday: function(a){ return BUILTINS.cdate.call(this,a).getDay()+1; },
  dateserial: function(a){ return new Date(Math.floor(toNum(a[0])), Math.floor(toNum(a[1]))-1, Math.floor(toNum(a[2]))); },
  dateadd: function(a){
    var iv = toStr(a[0]).toLowerCase(), n = Math.floor(toNum(a[1])), d = BUILTINS.cdate.call(this,[a[2]]);
    var r = new Date(d.getTime());
    if (iv === 'yyyy') r.setFullYear(r.getFullYear()+n);
    else if (iv === 'm') r.setMonth(r.getMonth()+n);
    else if (iv === 'd') r.setDate(r.getDate()+n);
    else if (iv === 'ww') r.setDate(r.getDate()+n*7);
    return r;
  },
  datediff: function(a){
    var iv = toStr(a[0]).toLowerCase(), d1 = BUILTINS.cdate.call(this,[a[1]]), d2 = BUILTINS.cdate.call(this,[a[2]]);
    if (iv === 'd') return Math.round((d2-d1)/DAY);
    if (iv === 'm') return (d2.getFullYear()-d1.getFullYear())*12 + (d2.getMonth()-d1.getMonth());
    if (iv === 'yyyy') return d2.getFullYear()-d1.getFullYear();
    return Math.round((d2-d1)/DAY);
  },
  chr: function(a){ return String.fromCharCode(Math.floor(toNum(a[0]))); },
  asc: function(a){ return toStr(a[0]).charCodeAt(0) || 0; },
  space: function(a){ return ' '.repeat(Math.max(0,Math.floor(toNum(a[0])))); },
  string: function(a){ return toStr(a[1]).charAt(0).repeat(Math.max(0,Math.floor(toNum(a[0])))); },
  strreverse: function(a){ return toStr(a[0]).split('').reverse().join(''); },
  hex: function(a){ return Math.floor(toNum(a[0])).toString(16).toUpperCase(); },
  rgb: function(a){ return Math.floor(toNum(a[0])) + Math.floor(toNum(a[1]))*256 + Math.floor(toNum(a[2]))*65536; },
  strconv: function(a){
    var s = toStr(a[0]), m = Math.floor(toNum(a[1]));
    if (m & 8) s = toHalf(s);
    if (m & 4) s = toFull(s);
    if (m === 1) s = s.toUpperCase();
    if (m === 2) s = s.toLowerCase();
    if (m === 3) s = s.replace(/\w\S*/g, function(x){ return x[0].toUpperCase()+x.slice(1).toLowerCase(); });
    return s;
  },
  iif: function(a){ return toBool(a[0]) ? a[1] : a[2]; },
  typename: function(a){
    var v = a[0];
    if (typeof v === 'number') return 'Double';
    if (typeof v === 'string') return 'String';
    if (typeof v === 'boolean') return 'Boolean';
    if (isDate(v)) return 'Date';
    if (v instanceof VbaArray) return 'Variant()';
    if (v === EMPTY || v === undefined) return 'Empty';
    if (v && v._t === 'Range') return 'Range';
    if (v && v._t === 'Sheet') return 'Worksheet';
    return 'Object';
  },
  nz: function(a){ return (a[0] === EMPTY || a[0] === undefined || a[0] === '') ? (a[1] === undefined ? 0 : a[1]) : a[0]; },
  doevents: function(){ return 0; },
  vartype: function(a){ return typeof a[0] === 'number' ? 5 : typeof a[0] === 'string' ? 8 : 0; }
};
function toHalf(s){
  return s.replace(/[！-～]/g, function(c){ return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); })
          .replace(/　/g,' ').replace(/‐|‑|‒|–|—|―|−|－/g,'-');
}
function toFull(s){ return s.replace(/[!-~]/g, function(c){ return String.fromCharCode(c.charCodeAt(0) + 0xFEE0); }).replace(/ /g,'　'); }

/* ---------- 公開 ---------- */
global.VBARuntime = {
  Sheet: Sheet, Workbook: Workbook, Interp: Interp,
  colLetter: colLetter, letterCol: letterCol, parseAddress: parseAddress,
  applyFormat: applyFormat, toStr: toStr, toNum: toNum, isDate: isDate,
  EMPTY: EMPTY, VbaArray: VbaArray, VbaError: VbaError,
  accessCount: function(){ return ACCESS.n; },
  run: function(src, workbook, opts){
    opts = opts || {};
    ACCESS.n = 0;
    var mod, it, name;
    try { mod = global.VBAParse(src); }
    catch (pe){ return { ok: false, out: [], error: pe && pe.isVbaError ? pe : new VbaError(0, String(pe && pe.message || pe)), phase: 'parse' }; }
    it = new Interp(mod, { workbook: workbook, inputs: opts.inputs, msgAnswer: opts.msgAnswer, maxSteps: opts.maxSteps });
    try { name = it.runModule(opts.entry); }
    catch (e){
      if (e && e.sig === 'end') return { ok: true, out: it.out, entry: null, ended: true, access: ACCESS.n };
      return { ok: false, out: it.out, access: ACCESS.n, error: e && e.isVbaError ? e : new VbaError(0, String(e && e.message || e)) };
    }
    return { ok: true, out: it.out, entry: name, access: ACCESS.n };
  }
};
})(window);
