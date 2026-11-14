/* ================================================================
 * 续航模型安全公式引擎（formula.js）
 * 纯词法 / 语法分析 + 递归下降求值，不使用 eval / new Function。
 *
 * 支持：
 *  - 数字字面量与四则运算 + - * /、幂 ^、括号、一元正负
 *  - 比较运算 > < >= <= = <>（结果为 1 / 0，供 IF 使用）
 *  - A1 单元格引用（B2、$B$2）与区间（B2:B13，仅限聚合函数参数）
 *  - 命名参数（Typ、可用容量、时长、电流、总耗电 等，由业务层注入）
 *  - 函数：SUM / AVG / MIN / MAX / ABS / ROUND / IF / POWER / SQRT（含中文别名）
 *  - 循环引用检测、错误传播（#REF! / #DIV/0! / #CYCLE! / #NAME? / #VALUE! / #SYNTAX!）
 * ================================================================ */
(function () {
  'use strict';

  var ERROR_TYPES = {
    REF: '#REF!',
    DIV0: '#DIV/0!',
    CYCLE: '#CYCLE!',
    NAME: '#NAME?',
    VALUE: '#VALUE!',
    SYNTAX: '#SYNTAX!',
  };

  function FError(type) { this.type = type; this.message = type; }
  FError.prototype.toString = function () { return this.type; };
  var ERROR_SET = {};
  Object.keys(ERROR_TYPES).forEach(function (k) { ERROR_SET[ERROR_TYPES[k]] = true; });
  // 同时识别 FError 实例与业务层 ctx 返回的鸭子类型错误对象 { type:'#REF!' }
  function isError(v) {
    return v instanceof FError || (v && typeof v.type === 'string' && ERROR_SET[v.type] === true);
  }
  function err(type) { return new FError(type); }

  // ---------- 列号工具：A <-> 1 ----------
  function colToLetter(col) {
    var s = '';
    var n = col;
    while (n > 0) {
      var m = (n - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }
  function letterToCol(letters) {
    var n = 0;
    for (var i = 0; i < letters.length; i += 1) {
      n = n * 26 + (letters.toUpperCase().charCodeAt(i) - 64);
    }
    return n;
  }
  function parseRef(text) {
    var m = /^\$?([A-Za-z]{1,3})\$?([0-9]{1,5})$/.exec(text);
    if (!m) return null;
    return { col: letterToCol(m[1]), row: parseInt(m[2], 10), label: m[1].toUpperCase() + m[2] };
  }

  // ---------- 词法分析 ----------
  function tokenize(input) {
    var s = String(input);
    var tokens = [];
    var i = 0;
    var isNameStart = function (ch) {
      return /[A-Za-z_\u4e00-\u9fa5]/.test(ch);
    };
    var isNamePart = function (ch) {
      return /[A-Za-z0-9_\u4e00-\u9fa5]/.test(ch);
    };
    while (i < s.length) {
      var ch = s[i];
      if (/\s/.test(ch)) { i += 1; continue; }
      if (/[0-9.]/.test(ch)) {
        var start = i;
        while (i < s.length && /[0-9.]/.test(s[i])) i += 1;
        if (i < s.length && (s[i] === 'e' || s[i] === 'E')) {
          i += 1;
          if (s[i] === '+' || s[i] === '-') i += 1;
          while (i < s.length && /[0-9]/.test(s[i])) i += 1;
        }
        var numText = s.slice(start, i);
        var num = Number(numText);
        if (!Number.isFinite(num)) throw err(ERROR_TYPES.SYNTAX);
        tokens.push({ kind: 'NUMBER', value: num });
        continue;
      }
      if (isNameStart(ch)) {
        var nameStart = i;
        while (i < s.length && isNamePart(s[i])) i += 1;
        // 混合引用 B$2：字母段后接 $ + 数字行号（$ 不在 isNamePart 内，单独吸收）
        if (s[i] === '$' && /[0-9]/.test(s[i + 1] || '')) {
          i += 1;
          while (i < s.length && /[0-9]/.test(s[i])) i += 1;
        }
        var word = s.slice(nameStart, i);
        if (/^\$?[A-Za-z]{1,3}\$?[0-9]{1,5}$/.test(word)) {
          tokens.push({ kind: 'REF', value: parseRef(word) });
        } else {
          tokens.push({ kind: 'NAME', value: word });
        }
        continue;
      }
      if (ch === '$') {
        // 绝对 / 混合引用：$B$2、$B2、B$2
        var refMatch = /^\$?[A-Za-z]{1,3}\$?[0-9]{1,5}/.exec(s.slice(i));
        if (!refMatch) throw err(ERROR_TYPES.SYNTAX);
        i += refMatch[0].length;
        tokens.push({ kind: 'REF', value: parseRef(refMatch[0]) });
        continue;
      }
      if (ch === '"' || ch === "'") {
        var quote = ch;
        var strStart = i + 1;
        i += 1;
        var str = '';
        while (i < s.length && s[i] !== quote) { str += s[i]; i += 1; }
        if (i >= s.length) throw err(ERROR_TYPES.SYNTAX);
        i += 1;
        tokens.push({ kind: 'STRING', value: str });
        continue;
      }
      var two = s.slice(i, i + 2);
      if (two === '>=' || two === '<=' || two === '<>') {
        tokens.push({ kind: 'OP', value: two }); i += 2; continue;
      }
      if ('+-*/^()><,:='.indexOf(ch) !== -1) {
        tokens.push({ kind: ch === '(' ? 'LPAREN' : ch === ')' ? 'RPAREN' : ch === ',' ? 'COMMA' : ch === ':' ? 'COLON' : 'OP', value: ch });
        i += 1; continue;
      }
      throw err(ERROR_TYPES.SYNTAX);
    }
    return tokens;
  }

  // ---------- 语法分析（递归下降） ----------
  // AST 节点：{ t:'num', v } / { t:'ref', ref } / { t:'range', a, b }
  //   { t:'name', v } / { t:'call', name, args } / { t:'bin', op, l, r } / { t:'neg', e } / { t:'str', v }
  function Parser(tokens) {
    this.tokens = tokens;
    this.pos = 0;
  }
  Parser.prototype.peek = function () { return this.tokens[this.pos] || null; };
  Parser.prototype.next = function () { return this.tokens[this.pos++]; };
  Parser.prototype.expectOp = function (op) {
    var tk = this.next();
    if (!tk || tk.kind !== 'OP' || tk.value !== op) throw err(ERROR_TYPES.SYNTAX);
  };
  Parser.prototype.expectRParen = function () {
    var tk = this.next();
    if (!tk || tk.kind !== 'RPAREN') throw err(ERROR_TYPES.SYNTAX);
  };
  Parser.prototype.parseExpression = function () { return this.parseComparison(); };
  Parser.prototype.parseComparison = function () {
    var left = this.parseAdditive();
    while (true) {
      var tk = this.peek();
      if (tk && tk.kind === 'OP' && ['>', '<', '>=', '<=', '=', '<>'].indexOf(tk.value) !== -1) {
        this.next();
        var right = this.parseAdditive();
        left = { t: 'bin', op: tk.value, l: left, r: right };
      } else break;
    }
    return left;
  };
  Parser.prototype.parseAdditive = function () {
    var left = this.parseMultiplicative();
    while (true) {
      var tk = this.peek();
      if (tk && tk.kind === 'OP' && (tk.value === '+' || tk.value === '-')) {
        this.next();
        left = { t: 'bin', op: tk.value, l: left, r: this.parseMultiplicative() };
      } else break;
    }
    return left;
  };
  Parser.prototype.parseMultiplicative = function () {
    var left = this.parseUnary();
    while (true) {
      var tk = this.peek();
      if (tk && tk.kind === 'OP' && (tk.value === '*' || tk.value === '/')) {
        this.next();
        left = { t: 'bin', op: tk.value, l: left, r: this.parseUnary() };
      } else break;
    }
    return left;
  };
  Parser.prototype.parseUnary = function () {
    var tk = this.peek();
    if (tk && tk.kind === 'OP' && (tk.value === '-' || tk.value === '+')) {
      this.next();
      var operand = this.parseUnary();
      return tk.value === '-' ? { t: 'neg', e: operand } : operand;
    }
    return this.parsePower();
  };
  Parser.prototype.parsePower = function () {
    var base = this.parsePrimary();
    var tk = this.peek();
    if (tk && tk.kind === 'OP' && tk.value === '^') {
      this.next();
      return { t: 'bin', op: '^', l: base, r: this.parseUnary() };
    }
    return base;
  };
  Parser.prototype.parsePrimary = function () {
    var tk = this.next();
    if (!tk) throw err(ERROR_TYPES.SYNTAX);
    if (tk.kind === 'NUMBER') return { t: 'num', v: tk.value };
    if (tk.kind === 'STRING') return { t: 'str', v: tk.value };
    if (tk.kind === 'REF') {
      var colon = this.peek();
      if (colon && colon.kind === 'COLON') {
        this.next();
        var end = this.next();
        if (!end || end.kind !== 'REF') throw err(ERROR_TYPES.SYNTAX);
        return { t: 'range', a: tk.value, b: end.value };
      }
      return { t: 'ref', ref: tk.value };
    }
    if (tk.kind === 'NAME') {
      var lp = this.peek();
      if (lp && lp.kind === 'LPAREN') {
        this.next();
        var args = [];
        if (this.peek() && this.peek().kind !== 'RPAREN') {
          args.push(this.parseExpression());
          while (this.peek() && this.peek().kind === 'COMMA') {
            this.next();
            args.push(this.parseExpression());
          }
        }
        var rp = this.next();
        if (!rp || rp.kind !== 'RPAREN') throw err(ERROR_TYPES.SYNTAX);
        return { t: 'call', name: tk.value, args: args };
      }
      return { t: 'name', v: tk.value };
    }
    if (tk.kind === 'LPAREN') {
      var node = this.parseExpression();
      this.expectRParen();
      // 括号后允许接区间（(B2):B4 少见，不支持）
      return node;
    }
    throw err(ERROR_TYPES.SYNTAX);
  };

  function parse(input) {
    var tokens = tokenize(input);
    if (!tokens.length) throw err(ERROR_TYPES.SYNTAX);
    var parser = new Parser(tokens);
    var ast = parser.parseExpression();
    if (parser.pos !== tokens.length) throw err(ERROR_TYPES.SYNTAX);
    return ast;
  }

  // ---------- 函数表 ----------
  var FN_ALIAS = {
    '求和': 'SUM', '平均值': 'AVG', '平均': 'AVG',
    '如果': 'IF', '四舍五入': 'ROUND', '绝对值': 'ABS',
    '最大值': 'MAX', '最小值': 'MIN', '开方': 'SQRT', '幂': 'POWER',
  };
  var AGG_FNS = { SUM: true, AVG: true, MIN: true, MAX: true };

  function flattenArgs(values) {
    var out = [];
    values.forEach(function (v) {
      if (v && v.__range) {
        v.refs.forEach(function (r) { out.push(r); });
      } else out.push(v);
    });
    return out;
  }

  // ---------- 求值 ----------
  // ctx: {
  //   getRef(ref): value | FError,              // ref = {col,row,label}
  //   getRange(a, b): { __range:true, refs:[...] } | FError,
  //   getName(name): value | FError,            // 数值 / 区间 / 字符串
  //   cellKey(ref): string                      // 循环检测用
  // }
  function evalAst(ast, ctx, visiting) {
    if (isError(ast)) return ast;
    switch (ast.t) {
      case 'num': return ast.v;
      case 'str': return ast.v;
      case 'neg': {
        var v = evalAst(ast.e, ctx, visiting);
        return isError(v) ? v : -v;
      }
      case 'ref': {
        var key = ctx.cellKey(ast.ref);
        if (visiting[key]) return err(ERROR_TYPES.CYCLE);
        return ctx.getRef(ast.ref);
      }
      case 'range': return ctx.getRange(ast.a, ast.b);
      case 'name': return ctx.getName(ast.v);
      case 'bin': return evalBin(ast, ctx, visiting);
      case 'call': return evalCall(ast, ctx, visiting);
      default: return err(ERROR_TYPES.SYNTAX);
    }
  }

  function evalBin(ast, ctx, visiting) {
    var op = ast.op;
    if (op === '^' || op === '*' || op === '/' || op === '+' || op === '-') {
      var l = evalAst(ast.l, ctx, visiting);
      if (isError(l)) return l;
      var r = evalAst(ast.r, ctx, visiting);
      if (isError(r)) return r;
      if (typeof l !== 'number' || typeof r !== 'number') return err(ERROR_TYPES.VALUE);
      switch (op) {
        case '+': return l + r;
        case '-': return l - r;
        case '*': return l * r;
        case '^': return Math.pow(l, r);
        case '/': return r === 0 ? err(ERROR_TYPES.DIV0) : l / r;
      }
    }
    // 比较运算
    var lc = evalAst(ast.l, ctx, visiting);
    if (isError(lc)) return lc;
    var rc = evalAst(ast.r, ctx, visiting);
    if (isError(rc)) return rc;
    if (typeof lc !== 'number' || typeof rc !== 'number') return err(ERROR_TYPES.VALUE);
    switch (op) {
      case '>': return lc > rc ? 1 : 0;
      case '<': return lc < rc ? 1 : 0;
      case '>=': return lc >= rc ? 1 : 0;
      case '<=': return lc <= rc ? 1 : 0;
      case '=': return lc === rc ? 1 : 0;
      case '<>': return lc !== rc ? 1 : 0;
    }
    return err(ERROR_TYPES.SYNTAX);
  }

  function evalCall(ast, ctx, visiting) {
    var fname = String(ast.name).toUpperCase();
    if (FN_ALIAS[ast.name]) fname = FN_ALIAS[ast.name];
    // IF 惰性求值：只计算条件与命中的分支，避免未命中分支里的错误被提前暴露
    if (fname === 'IF') {
      if (ast.args.length !== 3) return err(ERROR_TYPES.VALUE);
      var cond = evalAst(ast.args[0], ctx, visiting);
      if (isError(cond)) return cond;
      if (typeof cond !== 'number') return err(ERROR_TYPES.VALUE);
      return evalAst(ast.args[cond !== 0 ? 1 : 2], ctx, visiting);
    }
    var args = ast.args.map(function (a) { return evalAst(a, ctx, visiting); });
    for (var i = 0; i < args.length; i += 1) if (isError(args[i])) return args[i];

    if (AGG_FNS[fname]) {
      // 空单元格（null / undefined）在聚合中跳过；错误与非数值仍然中断
      var nums = flattenArgs(args).filter(function (v) { return v !== null && v !== undefined; });
      for (var j = 0; j < nums.length; j += 1) {
        if (isError(nums[j])) return nums[j];
        if (typeof nums[j] !== 'number') return err(ERROR_TYPES.VALUE);
      }
      if (!nums.length) return fname === 'SUM' ? 0 : err(ERROR_TYPES.VALUE);
      if (fname === 'SUM') return nums.reduce(function (a, b) { return a + b; }, 0);
      if (fname === 'AVG') return nums.reduce(function (a, b) { return a + b; }, 0) / nums.length;
      if (fname === 'MIN') return Math.min.apply(null, nums);
      if (fname === 'MAX') return Math.max.apply(null, nums);
    }
    switch (fname) {
      case 'ABS':
        if (args.length !== 1 || typeof args[0] !== 'number') return err(ERROR_TYPES.VALUE);
        return Math.abs(args[0]);
      case 'SQRT':
        if (args.length !== 1 || typeof args[0] !== 'number') return err(ERROR_TYPES.VALUE);
        if (args[0] < 0) return err(ERROR_TYPES.VALUE);
        return Math.sqrt(args[0]);
      case 'POWER':
        if (args.length !== 2 || typeof args[0] !== 'number' || typeof args[1] !== 'number') return err(ERROR_TYPES.VALUE);
        return Math.pow(args[0], args[1]);
      case 'ROUND':
        if (args.length < 1 || args.length > 2 || typeof args[0] !== 'number') return err(ERROR_TYPES.VALUE);
        var digits = args.length === 2 ? args[1] : 0;
        if (typeof digits !== 'number') return err(ERROR_TYPES.VALUE);
        var factor = Math.pow(10, digits);
        return Math.round(args[0] * factor) / factor;
      default:
        return err(ERROR_TYPES.NAME);
    }
  }

  // 对外：解析并求值一个公式字符串（不含开头的 =）
  function evaluate(formulaText, ctx) {
    try {
      var ast = parse(formulaText);
      var value = evalAst(ast, ctx, {});
      if (isError(value)) return { ok: false, error: value.type };
      if (typeof value !== 'number') return { ok: false, error: ERROR_TYPES.VALUE };
      if (!Number.isFinite(value)) return { ok: false, error: ERROR_TYPES.VALUE };
      return { ok: true, value: value };
    } catch (e) {
      if (e instanceof FError) return { ok: false, error: e.type };
      return { ok: false, error: ERROR_TYPES.SYNTAX };
    }
  }

  // 仅做语法检查（名称 / 引用是否存在交给业务 ctx）
  function checkSyntax(formulaText) {
    try {
      parse(formulaText);
      return { ok: true };
    } catch (e) {
      if (e instanceof FError) return { ok: false, error: e.type };
      return { ok: false, error: ERROR_TYPES.SYNTAX };
    }
  }

  window.ES_FORMULA = {
    evaluate: evaluate,
    checkSyntax: checkSyntax,
    parse: parse,
    colToLetter: colToLetter,
    letterToCol: letterToCol,
    parseRef: parseRef,
    ERROR_TYPES: ERROR_TYPES,
    isError: isError,
  };
})();
