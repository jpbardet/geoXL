/* ============================================================================================
 * geoXL — Excel custom functions translated from the GeoEXCEL LAMBDA library (v0.13).
 * Part 1: helpers that reproduce the Excel behaviour the LAMBDAs rely on.
 * ==========================================================================================*/
"use strict";

const PI = Math.PI;
const rad = (d) => (Number(d) * PI) / 180;
const deg = (r) => (Number(r) * 180) / PI;
const sin = Math.sin, cos = Math.cos, tan = Math.tan, atan = Math.atan, sqrt = Math.sqrt, abs = Math.abs;
const exp = Math.exp, ln = Math.log, log10 = Math.log10;
/** Excel ATAN2(x, y) — note the argument order (x first), unlike Math.atan2(y, x). */
const ATAN2 = (x, y) => Math.atan2(y, x);
/** Excel MOD: result has the sign of the divisor. */
const MOD = (a, b) => a - b * Math.floor(a / b);
const SIGN = (x) => (x > 0 ? 1 : x < 0 ? -1 : 0);
const ROUND = (x, n) => { const f = Math.pow(10, n); return Math.sign(x) * Math.round(Math.abs(x) * f) / f; };

/** Error raised inside a computation; the wrapper turns it into an Excel error value. */
class XlErr extends Error { constructor(code, msg) { super(msg || code); this.code = code; } }
const NA_CODE = "notAvailable", NUM_CODE = "invalidNumber", VALUE_CODE = "invalidValue", REF_CODE = "invalidReference", DIV_CODE = "divisionByZero";
function xlError(code) {
  if (typeof CustomFunctions !== "undefined" && CustomFunctions.Error) return new CustomFunctions.Error(CustomFunctions.ErrorCode[code]);
  return { xlError: code };            // used when testing in Node
}
/** Marker for #N/A inside returned arrays (e.g. Interpolate outside the data). */
const NAV = { na: true };

// ---------------------------------------------------------------- argument handling
const isMissing = (v) => v === undefined || v === null;
/** Number from an Excel cell value (blank -> 0, TRUE -> 1). */
function num(v) {
  if (Array.isArray(v)) v = v[0][0];
  if (v === "" || v === null || v === undefined || v === false) return 0;
  if (v === true) return 1;
  if (v === NAV) throw new XlErr(NA_CODE);
  const x = Number(v);
  if (Number.isNaN(x)) throw new XlErr(VALUE_CODE, "Number expected, got " + v);
  return x;
}
const str = (v) => { if (Array.isArray(v)) v = v[0][0]; return isMissing(v) ? "" : String(v); };
const lower = (v) => str(v).toLowerCase();
const upper = (v) => str(v).toUpperCase();
/** First value of an optional argument (undefined when omitted). */
function opt(v) {
  if (isMissing(v)) return undefined;
  if (Array.isArray(v)) { v = v[0] === undefined ? undefined : v[0][0]; }
  return v === null ? undefined : v;
}
/** Excel matrix -> 2-D array (scalars become 1x1). */
const mat = (v) => (Array.isArray(v) ? (Array.isArray(v[0]) ? v : [v]) : [[v]]);
/** Range -> column vector of numbers (row-major, like TOCOL). */
const colNum = (v) => mat(v).flat().map(num);
const colAny = (v) => mat(v).flat();
/** Column j (0-based) of a matrix. */
const column = (m, j) => mat(m).map((r) => r[j]);
const columnNum = (m, j) => column(m, j).map(num);
/** A point {x, y} given as 1x2 or 2x1 range. */
const point = (v) => { const f = colNum(v); return [f[0], f[1]]; };
/** Polygon / polyline given as two columns. */
const pts2 = (v) => mat(v).map((r) => [num(r[0]), num(r[1])]);

// ---------------------------------------------------------------- Excel-like lookups
/** XMATCH(key, list, 1) on text/number: exact match, else the next larger item; 0 if none (IFERROR -> 0). */
function xmatchNextLarger(key, list) {
  const i = list.indexOf(key);
  if (i >= 0) return i + 1;
  let best = -1;
  for (let k = 0; k < list.length; k++) if (list[k] > key && (best < 0 || list[k] < list[best])) best = k;
  return best + 1;
}
/** XMATCH(key, list, -1) linear: exact match, else the next smaller item; 0 if none. */
function xmatchNextSmaller(key, list) {
  const i = list.findIndex((v) => v === key);
  if (i >= 0) return i + 1;
  let best = -1;
  for (let k = 0; k < list.length; k++) if (list[k] < key && (best < 0 || list[k] > list[best])) best = k;
  return best + 1;
}
/** XMATCH(key, list, 1) numeric: exact or next larger; 0 if none. */
const xmatchNextLargerNum = (key, list) => xmatchNextLarger(key, list);
/** XLOOKUP(key, keys, values, ifNotFound, mode) with mode -1 (next smaller) or 1 (next larger). */
function xlookup(key, keys, values, ifNotFound, mode) {
  const i = mode === -1 ? xmatchNextSmaller(key, keys) : xmatchNextLarger(key, keys);
  return i ? values[i - 1] : ifNotFound;
}
/** FILTER(values, mask) — throws #CALC!-like error (#N/A here) when nothing matches and no default. */
function filter(values, mask, ifEmpty) {
  const out = values.filter((_, i) => mask[i]);
  if (out.length === 0) { if (ifEmpty !== undefined) return ifEmpty; throw new XlErr(NA_CODE, "FILTER: no match"); }
  return out;
}
const sum = (a) => a.reduce((s, v) => s + v, 0);
const average = (a) => { if (!a.length) throw new XlErr(DIV_CODE); return sum(a) / a.length; };
const geomean = (a) => { if (!a.length || a.some((v) => v <= 0)) throw new XlErr(NUM_CODE); return exp(sum(a.map(ln)) / a.length); };
const harmean = (a) => { if (!a.length || a.some((v) => v <= 0)) throw new XlErr(NUM_CODE); return a.length / sum(a.map((v) => 1 / v)); };
function slope(y, x) {
  const mx = average(x), my = average(y);
  let sxy = 0, sxx = 0;
  for (let i = 0; i < x.length; i++) { sxy += (x[i] - mx) * (y[i] - my); sxx += (x[i] - mx) ** 2; }
  if (sxx === 0) throw new XlErr(DIV_CODE);
  return sxy / sxx;
}
const intercept = (y, x) => average(y) - slope(y, x) * average(x);
/** AVERAGEIFS(values, keys, ">=" lo, keys, "<=" hi). */
function averageBetween(values, keys, lo, hi) {
  const sel = values.filter((_, i) => keys[i] >= lo && keys[i] <= hi);
  if (!sel.length) throw new XlErr(DIV_CODE, "AVERAGEIFS: no values in range");
  return average(sel);
}
const drop1 = (a) => a.slice(1), dropLast = (a) => a.slice(0, -1);
const pairMean = (a) => a.slice(1).map((v, i) => (v + a[i]) / 2);
const diffs = (a) => a.slice(1).map((v, i) => v - a[i]);
const cumsum = (a, start = 0) => { let s = start; return a.map((v) => (s += v)); };
const seq = (n, start = 1, step = 1) => Array.from({ length: n }, (_, i) => start + i * step);
/** SORT(UNIQUE(list)) for numbers (exact equality). */
const sortUnique = (a, desc) => { const u = [...new Set(a)]; u.sort((p, q) => (desc ? q - p : p - q)); return u; };
/** UNIQUE on rows (exact equality). */
function uniqueRows(rows) {
  const seen = new Set(), out = [];
  for (const r of rows) { const k = JSON.stringify(r); if (!seen.has(k)) { seen.add(k); out.push(r); } }
  return out;
}
/** Stable sort of rows by column j. */
const sortRows = (rows, j = 0, desc = false) =>
  rows.map((r, i) => [r, i]).sort((a, b) => (a[0][j] === b[0][j] ? a[1] - b[1] : (a[0][j] < b[0][j]) !== desc ? -1 : 1)).map((p) => p[0]);
/** Excel number-to-text in concatenation (15 significant digits, general format). */
const xlText = (v) => (typeof v === "number" ? String(Number(v.toPrecision(15))) : String(v));

/** Bessel function of the first kind, order 1 (Numerical Recipes rational approximation). */
function besselJ1(x) {
  const ax = abs(x);
  if (ax < 8.0) {
    const y = x * x;
    const a1 = x * (72362614232.0 + y * (-7895059235.0 + y * (242396853.1 + y * (-2972611.439 + y * (15704.48260 + y * (-30.16036606))))));
    const a2 = 144725228442.0 + y * (2300535178.0 + y * (18583304.74 + y * (99447.43394 + y * (376.9991397 + y * 1.0))));
    return a1 / a2;
  }
  const z = 8.0 / ax, y = z * z, xx = ax - 2.356194491;
  const a1 = 1.0 + y * (0.183105e-2 + y * (-0.3516396496e-4 + y * (0.2457520174e-5 + y * (-0.240337019e-6))));
  const a2 = 0.04687499995 + y * (-0.2002690873e-3 + y * (0.8449199096e-5 + y * (-0.88228987e-6 + y * 0.105787412e-6)));
  const ans = sqrt(0.636619772 / ax) * (cos(xx) * a1 - z * sin(xx) * a2);
  return x < 0 ? -ans : ans;
}

// ---------------------------------------------------------------- broadcasting (array lifting)
/** Shape of an Excel argument for broadcasting: [rows, cols]. Omitted -> [1,1]. */
const shapeOf = (v) => (isMissing(v) ? [1, 1] : [mat(v).length, mat(v)[0].length]);
function broadcastShape(args) {
  let r = 1, c = 1;
  for (const a of args) {
    const [ar, ac] = shapeOf(a);
    if (ar !== 1) { if (r !== 1 && r !== ar) throw new XlErr(VALUE_CODE, "Array sizes differ"); r = ar; }
    if (ac !== 1) { if (c !== 1 && c !== ac) throw new XlErr(VALUE_CODE, "Array sizes differ"); c = ac; }
  }
  return [r, c];
}
const pick = (v, i, j) => { if (isMissing(v)) return undefined; const m = mat(v); return m[m.length === 1 ? 0 : i][m[0].length === 1 ? 0 : j]; };
/**
 * Call a scalar function element by element, like Excel lifts a LAMBDA over arrays.
 * fixed: indices of arguments passed unchanged (tables, polygons).
 */
function lift(fn, args, fixed = []) {
  const liftArgs = args.filter((_, k) => !fixed.includes(k));
  const [R, C] = broadcastShape(liftArgs);
  const out = [];
  for (let i = 0; i < R; i++) {
    const row = [];
    for (let j = 0; j < C; j++) {
      const a = args.map((v, k) => (fixed.includes(k) ? v : pick(v, i, j)));
      try { row.push(fn(...a)); } catch (e) { row.push(e instanceof XlErr ? xlError(e.code) : xlError(VALUE_CODE)); }
    }
    out.push(row);
  }
  return out;
}
/** Like lift, but fn returns a row; element k of the inputs gives row k of the result. */
function liftRows(fn, args, fixed = []) {
  const liftArgs = args.filter((_, k) => !fixed.includes(k));
  const [R, C] = broadcastShape(liftArgs);
  const out = [];
  for (let i = 0; i < R; i++)
    for (let j = 0; j < C; j++) {
      const a = args.map((v, k) => (fixed.includes(k) ? v : pick(v, i, j)));
      try { out.push(fn(...a)); } catch (e) { out.push([e instanceof XlErr ? xlError(e.code) : xlError(VALUE_CODE)]); }
    }
  return out;
}

// ---------------------------------------------------------------- output
function cell(v) {
  if (v === NAV) return xlError(NA_CODE);
  if (typeof v === "number") return Number.isFinite(v) ? v : xlError(NUM_CODE);
  if (v === undefined || v === null) return "";
  return v;
}
/** Any result -> rectangular 2-D array for Excel (vectors become columns). */
function out(v) {
  let m;
  if (!Array.isArray(v)) m = [[v]];
  else if (v.length === 0) return [[xlError(NA_CODE)]];
  else if (!Array.isArray(v[0])) m = v.map((x) => [x]);
  else m = v;
  const w = Math.max(...m.map((r) => r.length));
  return m.map((r) => { const rr = r.map(cell); while (rr.length < w) rr.push(""); return rr; });
}
/** A single row -> 1 x n result. */
const row = (...vals) => [vals];
/** Run a computation and convert any internal error into an Excel error value. */
function safe(fnBody) {
  try { return out(fnBody()); }
  catch (e) { return [[xlError(e instanceof XlErr ? e.code : VALUE_CODE)]]; }
}
/** Same for lifted results (already 2-D). */
function safeLifted(fnBody) {
  try { return out(fnBody()); }
  catch (e) { return [[xlError(e instanceof XlErr ? e.code : VALUE_CODE)]]; }
}
