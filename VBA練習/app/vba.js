/* ================================================================
   VBA サブセット インタプリタ
   Sub/Function, Dim, If, For, For Each, Do, Select Case, With,
   On Error, 配列, Range/Cells/Font/Interior, 組み込み関数
   ================================================================ */
(function (global) {
'use strict';

var MAXROW = 1048576, MAXCOL = 16384;

/* ---------- 値 ---------- */
var EMPTY = { __empty: true };
var NOTHING = { __nothing: true };
function isEmptyVal(v){ return v === EMPTY || v === undefined || v === null; }

function VbaArray(dims, data){
  this.dims = dims;                       // [{lb,ub}, ...]
  var n = 1; for (var i=0;i<dims.length;i++) n *= (dims[i].ub - dims[i].lb + 1);
  this.data = data || new Array(n).fill(EMPTY);
}
VbaArray.prototype.flat = function(idx){
  if (idx.length !== this.dims.length) throw new VbaError(9, '添字の数が配列の次元と違います');
  var f = 0;
  for (var i=0;i<idx.length;i++){
    var d = this.dims[i], k = Math.floor(idx[i]);
    if (k < d.lb || k > d.ub) throw new VbaError(9, 'インデックスが有効範囲にありません（' + k + '）');
    f = f * (d.ub - d.lb + 1) + (k - d.lb);
  }
  return f;
};
VbaArray.prototype.get = function(idx){ return this.data[this.flat(idx)]; };
VbaArray.prototype.set = function(idx, v){ this.data[this.flat(idx)] = v; };

function VbaError(number, desc){ this.number = number; this.description = desc; this.isVbaError = true; }
VbaError.prototype.toString = function(){ return '実行時エラー ' + this.number + ': ' + this.description; };

function ErrValue(num){ this._t = 'ErrorValue'; this.number = num || 2042; }

/* ---------- 字句解析 ---------- */
var FULLWIDTH = { '（':'(', '）':')', '，':',', '．':'.', '＝':'=', '＋':'+', '－':'-', '＊':'*', '／':'/', '＜':'<', '＞':'>', '＆':'&', '：':':', '；':';' };

function isIdStart(ch){
  return /[A-Za-z_々〆ぁ-ゟ゠-ヿ㐀-鿿ｦ-ﾟＡ-Ｚａ-ｚ]/.test(ch);
}
function isIdChar(ch){ return isIdStart(ch) || /[0-9０-９]/.test(ch); }

function tokenize(src){
  var toks = [], i = 0, line = 1, n = src.length;
  function push(type, value){ toks.push({ type: type, value: value, line: line }); }
  while (i < n){
    var ch = src[i];
    if (ch === '\r'){ i++; continue; }
    if (ch === '\n'){ push('nl'); line++; i++; continue; }
    if (ch === ' ' || ch === '\t' || ch === '　'){ i++; continue; }
    // 行継続
    if (ch === '_' && !(i+1 < n && isIdChar(src[i+1]))){
      var j = i + 1;
      while (j < n && (src[j] === ' ' || src[j] === '\t' || src[j] === '\r')) j++;
      if (j >= n || src[j] === '\n'){ i = j + 1; line++; continue; }
    }
    if (ch === "'" || ch === '＇'){ while (i < n && src[i] !== '\n') i++; continue; }
    if (/[Rr]/.test(ch) && /^rem\b/i.test(src.slice(i, i+4)) && (toks.length === 0 || toks[toks.length-1].type === 'nl')){
      while (i < n && src[i] !== '\n') i++; continue;
    }
    if (ch === '"' || ch === '“' || ch === '”'){
      var s = '', close = '"'; i++;
      while (i < n){
        var c = src[i];
        if (c === '\n') throw new VbaError(0, (line) + '行目: 文字列が閉じられていません');
        if (c === '"' || c === '“' || c === '”'){
          if (src[i+1] === '"'){ s += '"'; i += 2; continue; }
          i++; break;
        }
        s += c; i++;
      }
      push('str', s); continue;
    }
    if (/[0-9]/.test(ch)){
      var st = i;
      while (i < n && /[0-9]/.test(src[i])) i++;
      if (src[i] === '.' && /[0-9]/.test(src[i+1] || '')){ i++; while (i < n && /[0-9]/.test(src[i])) i++; }
      if (/[eE]/.test(src[i] || '') && /[0-9+\-]/.test(src[i+1] || '')){ i += 2; while (i < n && /[0-9]/.test(src[i])) i++; }
      push('num', parseFloat(src.slice(st, i))); continue;
    }
    if (ch === '#'){ // 日付リテラル #2025/04/01#
      var st2 = ++i, s2 = '';
      while (i < n && src[i] !== '#' && src[i] !== '\n'){ s2 += src[i]; i++; }
      i++; push('date', s2.trim()); continue;
    }
    if (isIdStart(ch)){
      var st3 = i;
      while (i < n && isIdChar(src[i])) i++;
      var word = src.slice(st3, i);
      // 全角英数字を半角に寄せる
      word = word.replace(/[Ａ-Ｚａ-ｚ０-９]/g, function(c){ return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); });
      push('id', word); continue;
    }
    if (FULLWIDTH[ch]){
      throw new VbaError(0, line + '行目: 全角の記号「' + ch + '」が使われています。半角の「' + FULLWIDTH[ch] + '」に直してください');
    }
    var two = src.substr(i, 2);
    if (two === '<=' || two === '>=' || two === '<>' || two === ':='){ push('op', two); i += 2; continue; }
    if ('=+-*/\\^&<>(),.:'.indexOf(ch) >= 0){ push('op', ch); i++; continue; }
    throw new VbaError(0, line + '行目: 解釈できない文字「' + ch + '」があります');
  }
  push('nl'); push('eof');
  return toks;
}

/* ---------- 構文解析 ---------- */
function Parser(toks){ this.t = toks; this.p = 0; }
var P = Parser.prototype;
P.peek = function(k){ return this.t[this.p + (k||0)]; };
P.next = function(){ return this.t[this.p++]; };
P.line = function(){ return this.peek().line || 0; };
P.isId = function(w, k){ var tk = this.peek(k); return tk.type === 'id' && tk.value.toLowerCase() === w; };
P.isOp = function(w, k){ var tk = this.peek(k); return tk.type === 'op' && tk.value === w; };
P.eatId = function(w){ if (this.isId(w)){ this.next(); return true; } return false; };
P.expectOp = function(w){
  if (!this.isOp(w)) throw new VbaError(0, this.line() + '行目: 「' + w + '」が必要です');
  this.next();
};
P.skipNl = function(){ while (this.peek().type === 'nl') this.next(); };
P.endLine = function(){
  if (this.isOp(':')) { this.next(); return; }
  if (this.inlineMode) return;                 // 1行 If の中では改行を呼び出し元に残す
  if (this.peek().type === 'nl'){ this.next(); return; }
  if (this.peek().type === 'eof') return;
  throw new VbaError(0, this.line() + '行目: 文の終わりが正しくありません（「' + this.peek().value + '」）');
};

var TYPE_WORDS = ['long','integer','string','double','single','boolean','date','variant','currency','object','worksheet','workbook','range','collection','byte','decimal','xlcalculation'];

P.parseModule = function(){
  var mod = { procs: [], vars: [], consts: [], explicit: false };
  this.skipNl();
  while (this.peek().type !== 'eof'){
    if (this.isId('option')){
      if (this.isId('explicit', 1)) mod.explicit = true;
      while (this.peek().type !== 'nl' && this.peek().type !== 'eof') this.next();
      this.skipNl(); continue;
    }
    if (this.isId('private') || this.isId('public')) this.next();
    if (this.isId('sub') || this.isId('function')){ mod.procs.push(this.parseProc()); this.skipNl(); continue; }
    if (this.isId('dim')){ mod.vars.push(this.parseDim()); this.skipNl(); continue; }
    if (this.isId('const')){ mod.consts.push(this.parseConst()); this.skipNl(); continue; }
    if (this.peek().type === 'nl'){ this.next(); continue; }
    throw new VbaError(0, this.line() + '行目: モジュールの先頭に書けない文です（Sub の外に処理は書けません）');
  }
  return mod;
};

P.parseProc = function(){
  var kind = this.next().value.toLowerCase();       // sub | function
  var nameTok = this.next();
  if (nameTok.type !== 'id') throw new VbaError(0, nameTok.line + '行目: プロシージャ名がありません');
  var params = [];
  if (this.isOp('(')){
    this.next();
    while (!this.isOp(')')){
      var pm = { name: null, optional: false, def: null };
      if (this.isId('optional')){ this.next(); pm.optional = true; }
      if (this.isId('byval') || this.isId('byref')) this.next();
      pm.name = this.next().value;
      if (this.isId('as')){ this.next(); this.next(); }
      if (this.isOp('=')){ this.next(); pm.def = this.parseExpr(); }
      params.push(pm);
      if (this.isOp(',')) this.next(); else break;
    }
    this.expectOp(')');
  }
  if (this.isId('as')){ this.next(); this.next(); }
  this.endLine();
  var body = this.parseBlock(['end']);
  this.eatId('end'); this.next();                    // End Sub / End Function
  return { kind: kind, name: nameTok.value, params: params, body: body };
};

P.atBlockEnd = function(stops){
  var tk = this.peek();
  if (tk.type === 'eof') return true;
  if (tk.type !== 'id') return false;
  var w = tk.value.toLowerCase();
  for (var i=0;i<stops.length;i++) if (w === stops[i]) return true;
  return false;
};

P.parseBlock = function(stops){
  var out = [];
  this.skipNl();
  while (!this.atBlockEnd(stops)){
    if (this.peek().type === 'eof') break;
    var st = this.parseStatement();
    if (st) out.push(st);
    this.skipNl();
  }
  return out;
};

P.parseDim = function(){
  this.next();                                       // Dim
  var decls = [];
  while (true){
    var name = this.next().value;
    var dims = null;
    if (this.isOp('(')){
      this.next(); dims = [];
      while (!this.isOp(')')){
        var lo = this.parseExpr(), hi = null;
        if (this.isId('to')){ this.next(); hi = this.parseExpr(); }
        dims.push({ lo: lo, hi: hi });
        if (this.isOp(',')) this.next(); else break;
      }
      this.expectOp(')');
    }
    var type = 'Variant';
    if (this.isId('as')){ this.next(); if (this.isId('new')) this.next(); type = this.next().value; }
    decls.push({ name: name, dims: dims, type: type });
    if (this.isOp(',')) this.next(); else break;
  }
  return { type: 'dim', decls: decls, line: this.line() };
};

P.parseConst = function(){
  this.next();
  var name = this.next().value;
  if (this.isId('as')){ this.next(); this.next(); }
  this.expectOp('=');
  var v = this.parseExpr();
  return { type: 'const', name: name, value: v };
};

P.parseStatement = function(){
  var tk = this.peek(), ln = tk.line;
  if (tk.type === 'id'){
    var w = tk.value.toLowerCase();
    // ラベル
    if (this.isOp(':', 1) && !this.isOp('=', 2) && this.peek(2).type === 'nl'){
      var lab = this.next().value; this.next(); this.endLine();
      return { type: 'label', name: lab, line: ln };
    }
    switch (w){
      case 'dim': case 'static': { var d = this.parseDim(); this.endLine(); d.line = ln; return d; }
      case 'redim': return this.parseReDim();
      case 'const': { var c = this.parseConst(); this.endLine(); c.line = ln; return c; }
      case 'set': {
        this.next();
        var tgt = this.parseUnary();
        this.expectOp('=');
        var val = this.parseExpr(); this.endLine();
        return { type: 'assign', target: tgt, value: val, isSet: true, line: ln };
      }
      case 'if': return this.parseIf();
      case 'for': return this.parseFor();
      case 'do': return this.parseDo();
      case 'while': return this.parseWhile();
      case 'select': return this.parseSelect();
      case 'with': return this.parseWith();
      case 'exit': {
        this.next(); var what = this.next().value.toLowerCase(); this.endLine();
        return { type: 'exit', what: what, line: ln };
      }
      case 'call': {
        this.next(); var e = this.parseUnary(); this.endLine();
        return { type: 'exprstmt', expr: e, line: ln };
      }
      case 'on': return this.parseOnError();
      case 'resume': {
        this.next();
        var target = null;
        if (this.peek().type === 'id' && !this.isId('next')) target = this.next().value;
        else if (this.isId('next')) { this.next(); target = '__next'; }
        else if (this.peek().type === 'num') { this.next(); target = '0'; }
        this.endLine();
        return { type: 'resume', target: target, line: ln };
      }
      case 'goto': { this.next(); var lb = this.next().value; this.endLine(); return { type: 'goto', label: lb, line: ln }; }
      case 'end':
        if (this.peek(1).type === 'nl'){ this.next(); this.endLine(); return { type: 'endprog', line: ln }; }
        break;
      case 'erase': { this.next(); var ev = this.parseUnary(); this.endLine(); return { type: 'exprstmt', expr: ev, line: ln }; }
      case 'let': this.next(); break;
    }
  }
  // 代入 または Sub 呼び出し
  var lhs = this.parseUnary();
  if (this.isOp('=')){
    this.next();
    var rhs = this.parseExpr(); this.endLine();
    return { type: 'assign', target: lhs, value: rhs, line: ln };
  }
  // 括弧なしの引数
  var args = [];
  if (this.peek().type !== 'nl' && this.peek().type !== 'eof' && !this.isOp(':')){
    while (true){
      if (this.isOp(',')){ this.next(); args.push(null); continue; }
      args.push(this.parseArg());
      if (this.isOp(',')) this.next(); else break;
    }
  }
  this.endLine();
  if (args.length) return { type: 'callstmt', callee: lhs, args: args, line: ln };
  return { type: 'exprstmt', expr: lhs, line: ln };
};

P.parseArg = function(){
  if (this.peek().type === 'id' && this.isOp(':=', 1)){
    var nm = this.next().value; this.next();
    return { type: 'named', name: nm, value: this.parseExpr() };
  }
  return this.parseExpr();
};

P.parseReDim = function(){
  var ln = this.line(); this.next();
  var preserve = this.eatId('preserve');
  var name = this.next().value, dims = [];
  this.expectOp('(');
  while (!this.isOp(')')){
    var lo = this.parseExpr(), hi = null;
    if (this.isId('to')){ this.next(); hi = this.parseExpr(); }
    dims.push({ lo: lo, hi: hi });
    if (this.isOp(',')) this.next(); else break;
  }
  this.expectOp(')');
  if (this.isId('as')){ this.next(); this.next(); }
  this.endLine();
  return { type: 'redim', name: name, dims: dims, preserve: preserve, line: ln };
};

P.parseIf = function(){
  var ln = this.line(); this.next();
  var cond = this.parseExpr();
  if (!this.eatId('then')) throw new VbaError(0, ln + '行目: If に Then がありません');
  if (this.peek().type !== 'nl'){                    // 1行 If
    var thenStmt = this.parseInlineStatements();
    var elseStmt = [];
    if (this.isId('else')){ this.next(); elseStmt = this.parseInlineStatements(); }
    return { type: 'if', branches: [{ cond: cond, body: thenStmt }], elseBody: elseStmt, line: ln };
  }
  this.endLine();
  var branches = [{ cond: cond, body: this.parseBlock(['elseif','else','end']) }];
  var elseBody = [];
  while (this.isId('elseif')){
    this.next();
    var c2 = this.parseExpr();
    this.eatId('then'); this.endLine();
    branches.push({ cond: c2, body: this.parseBlock(['elseif','else','end']) });
  }
  if (this.isId('else')){ this.next(); this.endLine(); elseBody = this.parseBlock(['end']); }
  this.eatId('end'); this.eatId('if'); this.endLine();
  return { type: 'if', branches: branches, elseBody: elseBody, line: ln };
};

P.parseInlineStatements = function(){
  var out = [], saved = this.inlineMode;
  this.inlineMode = true;
  try {
    while (this.peek().type !== 'nl' && this.peek().type !== 'eof' && !this.isId('else')){
      out.push(this.parseStatement());
      if (this.isOp(':')) this.next();
    }
  } finally { this.inlineMode = saved; }
  return out;
};

P.parseFor = function(){
  var ln = this.line(); this.next();
  if (this.isId('each')){
    this.next();
    var v = this.next().value;
    if (!this.eatId('in')) throw new VbaError(0, ln + '行目: For Each に In がありません');
    var coll = this.parseExpr(); this.endLine();
    var body = this.parseBlock(['next']);
    this.eatId('next'); if (this.peek().type === 'id') this.next(); this.endLine();
    return { type: 'foreach', varName: v, coll: coll, body: body, line: ln };
  }
  var name = this.next().value;
  this.expectOp('=');
  var from = this.parseExpr();
  if (!this.eatId('to')) throw new VbaError(0, ln + '行目: For に To がありません');
  var to = this.parseExpr(), step = null;
  if (this.isId('step')){ this.next(); step = this.parseExpr(); }
  this.endLine();
  var body2 = this.parseBlock(['next']);
  this.eatId('next'); if (this.peek().type === 'id') this.next(); this.endLine();
  return { type: 'for', varName: name, from: from, to: to, step: step, body: body2, line: ln };
};

P.parseDo = function(){
  var ln = this.line(); this.next();
  var pre = null;
  if (this.isId('while')){ this.next(); pre = { until: false, cond: this.parseExpr() }; }
  else if (this.isId('until')){ this.next(); pre = { until: true, cond: this.parseExpr() }; }
  this.endLine();
  var body = this.parseBlock(['loop']);
  this.eatId('loop');
  var post = null;
  if (this.isId('while')){ this.next(); post = { until: false, cond: this.parseExpr() }; }
  else if (this.isId('until')){ this.next(); post = { until: true, cond: this.parseExpr() }; }
  this.endLine();
  return { type: 'do', pre: pre, post: post, body: body, line: ln };
};

P.parseWhile = function(){
  var ln = this.line(); this.next();
  var cond = this.parseExpr(); this.endLine();
  var body = this.parseBlock(['wend','loop']);
  this.next(); this.endLine();
  return { type: 'do', pre: { until: false, cond: cond }, post: null, body: body, line: ln };
};

P.parseSelect = function(){
  var ln = this.line(); this.next(); this.eatId('case');
  var subject = this.parseExpr(); this.endLine(); this.skipNl();
  var cases = [], elseBody = null;
  while (this.isId('case')){
    this.next();
    if (this.isId('else')){ this.next(); this.endLine(); elseBody = this.parseBlock(['case','end']); continue; }
    var tests = [];
    while (true){
      if (this.isId('is')){
        this.next();
        var op = this.next().value;
        tests.push({ kind: 'is', op: op, value: this.parseExpr() });
      } else {
        var a = this.parseExpr();
        if (this.isId('to')){ this.next(); tests.push({ kind: 'range', lo: a, hi: this.parseExpr() }); }
        else tests.push({ kind: 'eq', value: a });
      }
      if (this.isOp(',')) this.next(); else break;
    }
    this.endLine();
    cases.push({ tests: tests, body: this.parseBlock(['case','end']) });
  }
  this.eatId('end'); this.eatId('select'); this.endLine();
  return { type: 'select', subject: subject, cases: cases, elseBody: elseBody, line: ln };
};

P.parseWith = function(){
  var ln = this.line(); this.next();
  var obj = this.parseExpr(); this.endLine();
  var body = this.parseBlock(['end']);
  this.eatId('end'); this.eatId('with'); this.endLine();
  return { type: 'with', obj: obj, body: body, line: ln };
};

P.parseOnError = function(){
  var ln = this.line(); this.next();
  this.eatId('error');
  if (this.isId('resume')){ this.next(); this.eatId('next'); this.endLine(); return { type: 'onerror', mode: 'resume', line: ln }; }
  if (this.isId('goto')){
    this.next();
    var tk = this.next();
    this.endLine();
    if (tk.type === 'num' && tk.value === 0) return { type: 'onerror', mode: 'off', line: ln };
    return { type: 'onerror', mode: 'goto', label: tk.value, line: ln };
  }
  this.endLine();
  return { type: 'onerror', mode: 'off', line: ln };
};

/* --- 式 --- */
P.parseExpr = function(){ return this.parseOr(); };
P.parseOr = function(){
  var l = this.parseAnd();
  while (this.isId('or') || this.isId('xor')){ var o = this.next().value.toLowerCase(); l = { type: 'bin', op: o, l: l, r: this.parseAnd() }; }
  return l;
};
P.parseAnd = function(){
  var l = this.parseNot();
  while (this.isId('and')){ this.next(); l = { type: 'bin', op: 'and', l: l, r: this.parseNot() }; }
  return l;
};
P.parseNot = function(){
  if (this.isId('not')){ this.next(); return { type: 'un', op: 'not', e: this.parseNot() }; }
  return this.parseCmp();
};
P.parseCmp = function(){
  var l = this.parseConcat();
  while (true){
    if (this.peek().type === 'op' && ['=','<>','<','>','<=','>='].indexOf(this.peek().value) >= 0){
      var o = this.next().value;
      l = { type: 'bin', op: o, l: l, r: this.parseConcat() };
    } else if (this.isId('is')){
      this.next(); l = { type: 'bin', op: 'is', l: l, r: this.parseConcat() };
    } else if (this.isId('like')){
      this.next(); l = { type: 'bin', op: 'like', l: l, r: this.parseConcat() };
    } else break;
  }
  return l;
};
P.parseConcat = function(){
  var l = this.parseAdd();
  while (this.isOp('&')){ this.next(); l = { type: 'bin', op: '&', l: l, r: this.parseAdd() }; }
  return l;
};
P.parseAdd = function(){
  var l = this.parseMul();
  while (this.isOp('+') || this.isOp('-')){ var o = this.next().value; l = { type: 'bin', op: o, l: l, r: this.parseMul() }; }
  return l;
};
P.parseMul = function(){
  var l = this.parseIntDiv();
  while (this.isOp('*') || this.isOp('/') || this.isId('mod')){
    var o = this.next().value.toLowerCase();
    l = { type: 'bin', op: o, l: l, r: this.parseIntDiv() };
  }
  return l;
};
P.parseIntDiv = function(){
  var l = this.parsePow();
  while (this.isOp('\\')){ this.next(); l = { type: 'bin', op: '\\', l: l, r: this.parsePow() }; }
  return l;
};
P.parsePow = function(){
  var l = this.parseUnary();
  while (this.isOp('^')){ this.next(); l = { type: 'bin', op: '^', l: l, r: this.parseUnary() }; }
  return l;
};
P.parseUnary = function(){
  if (this.isOp('-')){ this.next(); return { type: 'un', op: '-', e: this.parseUnary() }; }
  if (this.isOp('+')){ this.next(); return this.parseUnary(); }
  return this.parsePostfix();
};
P.parsePostfix = function(){
  var e = this.parsePrimary();
  while (true){
    if (this.isOp('.')){
      this.next();
      var nameTok = this.next();
      if (nameTok.type !== 'id') throw new VbaError(0, nameTok.line + '行目: 「.」のあとにメンバー名がありません');
      var args = null;
      if (this.isOp('(')){ args = this.parseArgList(); }
      e = { type: 'member', target: e, name: nameTok.value, args: args, line: nameTok.line };
    } else if (this.isOp('(')){
      e = { type: 'call', callee: e, args: this.parseArgList(), line: this.line() };
    } else break;
  }
  return e;
};
P.parseArgList = function(){
  this.expectOp('(');
  var args = [];
  while (!this.isOp(')')){
    if (this.isOp(',')){ this.next(); args.push(null); continue; }
    args.push(this.parseArg());
    if (this.isOp(',')) this.next(); else break;
  }
  this.expectOp(')');
  return args;
};
P.parsePrimary = function(){
  var tk = this.peek();
  if (tk.type === 'num'){ this.next(); return { type: 'lit', value: tk.value }; }
  if (tk.type === 'str'){ this.next(); return { type: 'lit', value: tk.value }; }
  if (tk.type === 'date'){ this.next(); return { type: 'lit', value: new Date(tk.value.replace(/-/g,'/')) }; }
  if (tk.type === 'op' && tk.value === '('){
    this.next(); var e = this.parseExpr(); this.expectOp(')'); return e;
  }
  if (tk.type === 'op' && tk.value === '.'){       // With 内の .Foo
    this.next();
    var nm = this.next();
    var args = null;
    if (this.isOp('(')) args = this.parseArgList();
    return { type: 'member', target: { type: 'withref' }, name: nm.value, args: args, line: nm.line };
  }
  if (tk.type === 'id'){
    this.next();
    var w = tk.value.toLowerCase();
    if (w === 'true') return { type: 'lit', value: true };
    if (w === 'false') return { type: 'lit', value: false };
    if (w === 'nothing') return { type: 'lit', value: NOTHING };
    if (w === 'empty') return { type: 'lit', value: EMPTY };
    if (w === 'new'){ var t2 = this.next(); return { type: 'new', className: t2.value }; }
    return { type: 'id', name: tk.value, line: tk.line };
  }
  throw new VbaError(0, (tk.line||0) + '行目: 式が必要です（「' + (tk.value === undefined ? tk.type : tk.value) + '」）');
};

global.VBAParse = function(src){ return new Parser(tokenize(src)).parseModule(); };
global.VBAInternals = { EMPTY: EMPTY, NOTHING: NOTHING, VbaArray: VbaArray, VbaError: VbaError, ErrValue: ErrValue, MAXROW: MAXROW, MAXCOL: MAXCOL, isEmptyVal: isEmptyVal };
})(window);
