/* ============================================================
 * 华哥报价计算器 手机版 —— 计算引擎 V1.0
 * 与电脑版 main.py compute() 逐式对齐（含 V4.7 圆棒 π/4、
 * V4.8 自动车床产能、V4.9 加工中心费与调机费切换、V5.2 默认值）
 * Excel 语义算术：错误值传播、分母 0 → "—"（后续按 0 参与运算）
 * ============================================================ */
"use strict";

var PI = Math.PI;
var ERR = "#VALUE!";   // Excel 语义错误值
var DASH = "—";        // 分母为 0 的空占位（参与后续运算时按 0 处理）

function isErr(v) { return typeof v === "string" && v.charCodeAt(0) === 35; }

function coerce(v) {
  if (v === null || v === undefined) return 0;
  if (typeof v === "boolean") return v ? 1 : 0;
  if (typeof v === "number") return isFinite(v) ? v : ERR;
  if (typeof v !== "string") return ERR;
  var s = v.trim();
  if (s === "" || s === "—" || s === "-") return 0;
  if (/^[+-]?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?$/.test(s)) return parseFloat(s);
  return ERR;
}

function _num2(fn) {
  return function (a, b) {
    a = isErr(a) ? a : coerce(a);
    b = isErr(b) ? b : coerce(b);
    if (isErr(a)) return a;
    if (isErr(b)) return b;
    return fn(a, b);
  };
}
var xadd = _num2(function (a, b) { return a + b; });
var xsub = _num2(function (a, b) { return a - b; });
var xmul = _num2(function (a, b) { return a * b; });

function xdiv(a, b) {
  a = isErr(a) ? a : coerce(a);
  b = isErr(b) ? b : coerce(b);
  if (isErr(a)) return a;
  if (isErr(b)) return b;
  if (b === 0) return DASH;
  return a / b;
}
function xneg(a) { a = coerce(a); return isErr(a) ? a : -a; }
function xint(a) { a = coerce(a); return isErr(a) ? a : Math.floor(a); }

/* 截面积 mm²：圆棒 πr²；六角 0.866×对边距²；管料 π/4×(外径²−内径²) */
function area_of(b, shape, inner) {
  if (isErr(b)) return b;
  if (shape === "管料") {
    var c = coerce(inner);
    if (isErr(c) || c <= 0 || c >= b) return ERR;
    return PI * (b * b - c * c) / 4;
  }
  return shape === "圆棒" ? PI * Math.pow(b / 2, 2) : 0.866 * b * b;
}

/* ==================== 核心计算 ====================
 * inp: 键值对象（B3~B56、CAP_L、CAP_S1~4、DAY_RATE、_3D_WEIGHT）
 * 返回：结果键值对象（B16~B54/D32/E40/LATHE_FEE/...） */
function compute(inp) {
  function g(k, d) {
    if (inp && Object.prototype.hasOwnProperty.call(inp, k) && inp[k] !== undefined) return inp[k];
    return d === undefined ? null : d;
  }
  var B4 = g("B4", "");
  var B5 = g("B5", "");
  var B6 = coerce(g("B6")), B7 = coerce(g("B7"));
  var B6B = g("B6B"); // 管料内径（仅毛坯=管料时参与计算）
  var B8 = coerce(g("B8")), B9 = coerce(g("B9")), B10 = coerce(g("B10"));
  var B11 = coerce(g("B11"));
  var raw13 = g("B13");
  var r13 = (raw13 === null || raw13 === undefined || String(raw13).trim() === "") ? ERR : coerce(raw13);
  var raw13b = g("B13_MC");
  var r13b = (raw13b === null || raw13b === undefined || String(raw13b).trim() === "") ? ERR : coerce(raw13b);

  var out = {};
  // V4.8：自动车床切刀余量 2.5mm，其它设备 3mm
  out["B16"] = xadd(B7, B4 === "自动车床" ? 2.5 : 3.0);
  if (B4 === "走心机") out["B17"] = 2260.0;
  else if (B4 === "走刀机") out["B17"] = 2370.0;
  else if (B4 === "自动车床") out["B17"] = 2440.0; // 2500-60（料尾）
  else out["B17"] = DASH; // 加工中心

  out["B18"] = area_of(B6, B5, B6B);
  out["B19"] = B4 === "加工中心" ? DASH : xdiv(xmul(xmul(out["B18"], 2500.0), r13), 1e6);
  out["B20"] = B4 === "加工中心" ? DASH : xint(xdiv(out["B17"], out["B16"]));
  out["B21"] = xmul(xmul(B8, B9), B10);
  out["B22"] = B4 === "加工中心" ? xdiv(xmul(out["B21"], r13b), 1e6) : DASH;
  out["B23"] = B4 === "加工中心" ? out["B22"] : xdiv(out["B19"], out["B20"]);
  out["B24"] = xmul(out["B23"], B11);

  var r_active = B4 === "加工中心" ? r13b : r13;
  var ADD_ROWS = [28, 29, 30, 31, 101, 102, 103, 104];
  for (var i = 0; i < ADD_ROWS.length; i++) {
    var rr = ADD_ROWS[i];
    var d = coerce(g("B" + rr));
    var length = coerce(g("C" + rr));
    // V4.7：圆棒按真实圆棒料截面积 π/4×D² 计算（与主计算区圆棒 B18 口径一致）
    out["D" + rr] = xdiv(xmul(xmul(xmul(xmul(d, d), length), r_active), PI / 4), 1e6);
  }
  var total = out["D28"];
  for (var j = 1; j < ADD_ROWS.length; j++) {
    var r2 = ADD_ROWS[j];
    var term = out["D" + r2];
    total = xadd(total, g("E" + r2) === "减" ? xneg(term) : term);
  }
  out["D32"] = total;

  for (var k = 0; k < 3; k++) {
    var r3 = 37 + k;
    out["E" + r3] = xdiv(xmul(xmul(xmul(coerce(g("B" + r3)), coerce(g("C" + r3))),
      coerce(g("D" + r3))), r13b), 1e6);
  }
  var RB_ROWS = [107, 108];
  for (var m = 0; m < RB_ROWS.length; m++) {
    var rr4 = RB_ROWS[m];
    var d4 = coerce(g("B" + rr4));
    var len4 = coerce(g("C" + rr4));
    // V4.7：圆棒按 π/4×D²；方料 E37~E39 仍为 W×H×L 不变
    out["E" + rr4] = xdiv(xmul(xmul(xmul(xmul(d4, d4), len4), r_active), PI / 4), 1e6);
  }
  total = out["E37"];
  for (var n = 1; n < 3; n++) {
    var r5 = 37 + n;
    var term5 = out["E" + r5];
    total = xadd(total, g("F" + r5) === "减" ? xneg(term5) : term5);
  }
  // 圆棒1/圆棒2 都按 F107/F108 的加减方向计入
  total = xadd(total, g("F107") === "减" ? xneg(out["E107"]) : out["E107"]);
  total = xadd(total, g("F108") === "减" ? xneg(out["E108"]) : out["E108"]);
  out["E40"] = total;

  /* ==================== 成本核算 ==================== */
  var B44 = coerce(g("B44")), B45 = coerce(g("B45"));
  var B48 = coerce(g("B48")), B51 = coerce(g("B51"));
  var B56 = coerce(g("B56")); // V4.9：加工中心加工秒数

  // 报价表自算净重：加工中心用 E40，走心/走刀用 D32
  var net_quote = B4 === "加工中心" ? out["E40"] : out["D32"];

  // 3D 净重注入点（手机版无 3D，恒为空 → 用报价表自算净重）
  var _3d_raw = g("_3D_WEIGHT");
  var _3d_set = _3d_raw !== null && _3d_raw !== undefined && String(_3d_raw).trim() !== "";
  var net = _3d_set ? coerce(_3d_raw) : net_quote;

  out["_NET_SOURCE"] = _3d_set ? "3D" : (B4 === "加工中心" ? "E40" : "D32");
  out["_NET_USED"] = net;
  out["_NET_QUOTE"] = net_quote;

  // B47 废料抵扣 = (毛重 B23 − 当前净重) × 废料单价
  out["B47"] = xmul(xsub(out["B23"], net), B45);
  // B46 材料费：毛重始终为报价表毛重 B23
  out["B46"] = xadd(xsub(xmul(xmul(out["B23"], 1.03), B44), out["B47"]), xdiv(100.0, B11));
  // B49 表面处理费 = 当前净重 × 单价 + 200/数量
  out["B49"] = xadd(xmul(net, B48), xdiv(200.0, B11));
  if (B4 === "自动车床") {
    // V4.8：自动车床按产能计算加工费，不按加工秒数
    out["B50"] = 0.0;
    out["B52"] = DASH;
    out["B56FEE"] = 0.0;
    var day_rate = coerce(g("DAY_RATE"));
    var cap_l = coerce(g("CAP_L"));
    var lathe_fee = (cap_l && cap_l > 0) ? xmul(xadd(xdiv(B11, cap_l), 0.5), xdiv(day_rate, B11)) : 0.0;
    var sec_total = 0.0;
    for (var s = 1; s <= 4; s++) {
      var cap_s = coerce(g("CAP_S" + s));
      if (cap_s && cap_s > 0) {
        var sf = xmul(xadd(xdiv(B11, cap_s), 0.2), xdiv(day_rate, B11));
        out["SEC_FEE_" + s] = sf;
        sec_total = xadd(sec_total, sf);
      } else {
        out["SEC_FEE_" + s] = 0.0;
      }
    }
    out["LATHE_FEE"] = lathe_fee;
    out["SEC_FEE_TOTAL"] = sec_total;
    out["B53"] = xadd(lathe_fee, sec_total);
  } else {
    // V4.9 调机费：任填一个秒数=600，两项都填=1200
    var has_lathe = typeof B51 === "number" && B51 > 0;
    var has_mc = typeof B56 === "number" && B56 > 0;
    out["B50"] = xdiv((has_lathe && has_mc) ? 1200.0 : 600.0, B11);
    out["B52"] = B4 === "走刀机" ? 50.0 : 60.0;
    out["B53"] = xmul(xdiv(B51, 3600.0), out["B52"]);
    // V4.9：加工中心费 = 加工中心秒数÷3600×60元/h
    out["B56FEE"] = has_mc ? xmul(xdiv(B56, 3600.0), 60.0) : 0.0;
  }
  // B55 清洗费（单件）：净重×3元/kg，总清洗费保底 100 元
  var per_clean = xmul(net, 3.0);
  var per_floor = xdiv(100.0, B11);
  if (!(isErr(per_clean) || isErr(per_floor)) && per_clean < per_floor) {
    out["B55"] = per_floor;
  } else {
    out["B55"] = per_clean;
  }
  // V4.9：总成本 = 材料费+表面处理费+调机费+加工费+加工中心费+清洗费
  out["B54"] = xadd(xadd(xadd(xadd(xadd(out["B46"], out["B49"]), out["B50"]),
    out["B53"]), out["B56FEE"] !== undefined ? out["B56FEE"] : 0.0), out["B55"]);
  return out;
}

/* ==================== 显示格式（与电脑版一致） ==================== */
function fgeneral(v) {
  if (isErr(v) || typeof v === "string") return String(v);
  if (Number.isInteger(v) && Math.abs(v) < 1e15) return String(v);
  return String(parseFloat(v.toPrecision(11)));
}
function ffixed(v, n) {
  if (isErr(v) || typeof v === "string") return String(v);
  return v.toFixed(n);
}

/* ==================== 材料与默认值（与电脑版一致） ==================== */
var MATERIALS = [
  ["2024铝", 2.7], ["6061铝", 2.73], ["7075铝", 2.81], ["5052铝", 2.68],
  ["黄铜H62", 8.43], ["黄铜H68", 8.5], ["黄铜H59", 8.4],
  ["磷青铜QSn6.5-0.1", 8.8], ["铝青铜QA19-4", 7.46], ["白铜B19", 8.9],
  ["无氧铜C1020", 8.94], ["紫铜T2/T3", 8.9],
  ["不锈钢304", 7.93], ["不锈钢316", 7.98],
  ["碳钢Q235/45#", 7.85], ["合金钢40Cr", 7.85], ["铸铁HT200", 7.25],
  ["钛合金TC4", 4.51], ["锌合金ZA", 6.7], ["镁合金AZ31", 1.78], ["尼龙PA66", 1.15]
];
var EQUIPMENTS = ["走心机", "走刀机", "加工中心", "自动车床"];
var BLANKS = ["圆棒", "六角", "管料"];
var SIGNS = ["加", "减"];

var DEFAULTS = {
  "B3": "",
  "B4": "走心机", "B5": "圆棒",
  "B6": "", "B6B": "", "B7": "", "B8": "", "B9": "", "B10": "",
  "B11": "1",
  "B12": "2024铝", "B13": "2.7", "B13_MC": "2.7",
  "B28": "", "C28": "", "E28": "加",
  "B29": "", "C29": "", "E29": "减",
  "B30": "0", "C30": "0", "E30": "加",
  "B31": "0", "C31": "0", "E31": "加",
  "B101": "0", "C101": "0", "E101": "加",
  "B102": "0", "C102": "0", "E102": "加",
  "B103": "0", "C103": "0", "E103": "加",
  "B104": "0", "C104": "0", "E104": "加",
  "B37": "0", "C37": "0", "D37": "0", "F37": "加",
  "B38": "0", "C38": "0", "D38": "0", "F38": "减",
  "B39": "0", "C39": "0", "D39": "0", "F39": "加",
  "B107": "0", "C107": "0", "F107": "加",
  "B108": "0", "C108": "0", "F108": "减",
  "B44": "0", "B45": "0", "B48": "0", "B51": "0", "B56": "0",
  "CAP_L": "0", "CAP_S1": "0", "CAP_S2": "0", "CAP_S3": "0", "CAP_S4": "0",
  "DAY_RATE": "300",
  "_3D_WEIGHT": ""
};

/* 导出（浏览器 + Node 测试两用） */
if (typeof module !== "undefined" && module.exports) {
  module.exports = { PI: PI, ERR: ERR, DASH: DASH, isErr: isErr, coerce: coerce,
    xadd: xadd, xsub: xsub, xmul: xmul, xdiv: xdiv, xneg: xneg, xint: xint,
    area_of: area_of, compute: compute, fgeneral: fgeneral, ffixed: ffixed,
    MATERIALS: MATERIALS, EQUIPMENTS: EQUIPMENTS, BLANKS: BLANKS, SIGNS: SIGNS,
    DEFAULTS: DEFAULTS };
}
