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
/* ============================================================================================
 * Part 2: engineering functions (plain JavaScript values). Each follows the LAMBDA of the
 * same name in GeoEXCEL v0.13 line by line; comments flag behaviour kept from the original.
 * ==========================================================================================*/

const G = {};   // namespace of internal implementations

G.version = () => "geoXL 1.0 (JavaScript add-in, from GeoEXCEL 0.13) 2026-10-17";

// ---------------------------------------------------------------- bearing capacity
const AUTH6 = ["terzaghi", "prandtl", "meyerhof", "hansen", "vesic", "ec7"];
G.bearing_factors = (phi, author) => {
  const f = rad(phi), a = xmatchNextLarger(lower(author), AUTH6);
  const Nq = a === 1 ? exp((3 * PI / 2 - f) * tan(f)) / 2 / cos(PI / 4 + f / 2) ** 2
    : a > 1 ? exp(PI * tan(f)) * tan(PI / 4 + f / 2) ** 2 : 0;
  const Nc = a === 1 && phi === 0 ? 1 + 3 * PI / 2 : a > 1 && phi === 0 ? 2 + PI : (Nq - 1) / tan(f);
  const Ng = { 1: 2 * (Nq + 1) * tan(f) / (1 + 0.4 * sin(f)), 2: 0, 3: (Nq - 1) * tan(1.4 * f), 4: 1.5 * (Nq - 1) * tan(f),
    5: 2 * (Nq + 1) * tan(f), 6: 2 * (Nq - 1) * tan(f), 0: "Error" }[a];
  return [Nc, Nq, Ng];
};
G.bearing_depthFactors = (DB, phi, author) => {
  const f = rad(phi), a = xmatchNextLarger(lower(author), AUTH6);
  const k = DB > 1 ? atan(DB) : DB;
  const dq = a === 3 ? (phi < 10 ? 1 : 1 + 0.1 * DB * tan(PI / 4 + f / 2)) : a === 4 || a === 5 ? 1 + 2 * k * tan(f) * (1 - sin(f)) ** 2 : 1;
  const dc = a === 3 ? 1 + 0.2 * DB * tan(PI / 4 + f / 2) : a === 4 || a === 5 ? 1 + 0.4 * k : 1;
  const dg = a === 3 ? dq : 1;
  return [dc, dq, dg];
};
G.bearing_inclinationFactors = (H, V, BL, phi, coh, author, A_) => {
  const f = rad(phi), auth = xmatchNextLarger(lower(author), AUTH6);
  const ach = coh * (A_ === undefined ? 1 : num(A_));
  const ac = ach > H ? H : ach;
  const t = V === 0 ? 0 : atan(H / V);
  const m = (2 + BL) / (1 + BL), tt = abs(H / (V + ac));
  const [Nc, Nq] = G.bearing_factors(phi, author);
  const ratio = () => (1 - (t / f > 1 ? 1 : t / f)) ** 2;
  const iq = auth === 3 ? (phi === 0 ? 0 : ratio()) : auth === 4 ? abs(1 - tt / 2) ** 3 : auth === 5 ? abs(1 - tt) ** m : auth === 6 ? 1 - t : 1;
  const ig = auth === 3 ? (phi === 0 ? 0 : ratio()) : auth === 4 ? abs(1 - 0.7 * tt) ** 3 : auth === 5 ? abs(1 - tt) ** m : auth === 6 ? (V === 0 ? 1 : 1 - H / V) : 1;
  const ic = auth === 3 ? iq
    : auth === 4 ? (phi === 0 ? 0.5 - sqrt(1 - H / ac) : (iq * Nq - 1) / (Nq - 1))
    : auth === 5 ? (phi === 0 ? 1 - H * m / ac / Nc : (iq * Nq - 1) / (Nq - 1))
    : auth === 6 ? (phi === 0 ? 0.5 + 0.5 * sqrt(1 - H / ac) : (iq * Nq - 1) / (Nq - 1)) : 1;
  return [Math.max(ic, 0), Math.max(iq, 0), Math.max(ig, 0)];
};
G.bearing_shapeFactors = (BL, phi, author, opt_) => {
  const optOmitted = opt_ === undefined;
  const f = rad(phi), auth = xmatchNextLarger(lower(author), AUTH6);
  const [Nc, Nq] = G.bearing_factors(phi, author);
  const Kp = tan(PI / 4 + f / 2) ** 2;
  const sq = auth === 3 ? (phi < 0 ? 1 : 1 + 0.1 * BL * Kp) : auth === 4 ? 1 + BL * sin(f) : auth === 5 ? 1 + BL * tan(f) : 1;
  const sc = auth === 1 ? 1 + 0.3 * BL : auth === 3 ? 1 + 0.2 * BL * Kp : auth === 4 || auth === 5 ? (phi === 0 ? 0.2 * BL : 1 + BL * Nq / Nc) : 1;
  const sg = auth === 1 ? (!optOmitted ? 0.6 : 1 - 0.2 * BL) : auth === 3 ? sq : auth === 4 || auth === 5 ? Math.max(0.6, 1 - 0.4 * BL) : 1;
  return [sc, sq, sg];
};
G.bearing_capacity = (coh, phi, q, gam, B, L, D, H, V, author) => {
  const BL = B / L, DB = D / B;
  const load = [coh, q, gam * B / 2];
  const s = G.bearing_shapeFactors(BL, phi, author), d = G.bearing_depthFactors(DB, phi, author);
  const i = G.bearing_inclinationFactors(H, V, BL, phi, coh, author), N = G.bearing_factors(phi, author);
  const nz = (v) => (typeof v === "number" ? v : 0);              // SUMPRODUCT treats text as 0
  return sum([0, 1, 2].map((k) => nz(s[k]) * nz(d[k]) * nz(i[k]) * nz(N[k]) * nz(load[k])));
};
G.bearing_Nq = (phi, author) => G.bearing_factors(phi, author)[1];
G.bearing_Skempton = (DB, BL) => {
  const x = DB > 6 ? 6 : DB;
  const Nc1 = -0.0027 * x ** 4 + 0.0607 * x ** 3 - 0.5116 * x ** 2 + 1.9247 * x + 6.2304;
  const Nc0 = -0.0021 * x ** 4 + 0.0504 * x ** 3 - 0.4354 * x ** 2 + 1.6433 * x + PI + 2;
  return Nc0 + (Nc1 - Nc0) * BL;
};

// ---------------------------------------------------------------- classification
G.classify_AASHTO = (pf10, pf40, pf200, LL_, PL_) => {
  const LL = LL_ === undefined ? 0 : num(LL_), PL = PL_ === undefined ? 0 : num(PL_), PIx = LL - PL;
  const g = Math.max(0, (pf200 - 35) * (0.2 + 0.005 * (LL - 40)) + 0.01 * (pf200 - 15) * (PIx - 10));
  const GI = "(" + xlText(g) + ")";
  if (pf10 < pf40) return "% passing No.10 must be larger than % passing No.40";
  if (pf40 < pf200) return "% passing No.40 must be larger than % passing No.200";
  if (LL < PL) return "liquid limit must be larger or equal plastic limit";
  if (pf10 < 50 && pf40 < 30 && pf200 < 15 && PIx < 6) return "A-1a - Stone fragment gravel and sand";
  if (pf40 < 50 && pf200 < 25 && PIx < 6) return "A-1b - Stone fragment gravel and sand";
  if (pf40 > 51 && pf200 < 10 && PL <= 0) return "A-3 - Fine sand";
  if (pf200 < 35 && LL < 40 && PIx < 10) return "A-2-4" + GI + " - Silty or clayey gravel sand";
  if (pf200 < 35 && LL > 41 && PIx < 10) return "A-2-5" + GI + " - Silty or clayey gravel sand";
  if (pf200 < 35 && LL < 40 && PIx > 11) return "A-2-6" + GI + " - Silty or clayey gravel sand";
  if (pf200 < 35 && LL > 41 && PIx > 11) return "A-2-7" + GI + " - Silty or clayey gravel sand";
  if (pf200 >= 36 && LL < 40 && PIx < 10) return "A-4" + GI + " - Silty soil";
  if (pf200 >= 36 && LL > 41 && PIx < 10) return "A-5" + GI + " - Silty soil";
  if (pf200 >= 36 && LL < 40 && PIx > 11) return "A-6" + GI + " - Clayey soil";
  if (pf200 >= 36 && LL > 41 && PIx > 11 && PL > 30) return "A-7-5" + GI + " - Clayey soil";
  if (pf200 >= 36 && LL > 41 && PIx > 11 && PL < 30) return "A-7-6" + GI + " - Clayey soil";
  return "N-C";
};
G.classify_Plasticity = (LL, PL, organic) => {
  const org = organic === undefined ? 0 : num(organic), PIx = LL - PL;
  const aboveU = LL < 7 ? PIx > LL : LL <= 15 ? PIx > 7 : PIx > 0.9 * (LL - 8);
  const aboveA = LL < 25.48 ? PIx > 4 : PIx > 0.73 * (LL - 20);
  if (LL < PL) return "Error: Liquid limit must be larger or equal to plastic limit";
  if (aboveU) return "Error: Above U line in plasticity chart";
  if (4 <= PIx && PIx <= 7 && LL > PIx && PIx >= 0.73 * (LL - 20)) return "CL & ML";
  if (aboveA) return (org === 0 ? "C" : "O") + (LL > 50 ? "H" : "L");
  return (org === 0 ? "M" : "O") + (LL > 50 ? "H" : "L");
};
G.classify_USC = (pf4, pf200, Cu_, Cc_, LL_, PL_, org_) => {
  const o = (v) => (v === undefined ? 0 : num(v));
  const organic = o(org_), Cu = o(Cu_), Cc = o(Cc_), LL = o(LL_), PL = o(PL_);
  const fines = G.classify_Plasticity(LL, PL);           // organic not passed (as in the LAMBDA)
  const clayFines = fines.slice(0, 1) === "C";
  const isSand = pf4 - pf200 > 100 - pf4;
  const wg = isSand ? Cu > 6 && Cc > 1 && Cc < 3 : Cu > 4 && Cc > 1 && Cc < 3;
  if (pf4 < pf200) return "Error: Percent passing No.4 must be larger than percent passing No.200";
  if (organic > 90) return "Pt";
  if (fines.slice(0, 3) === "Err") return fines;
  if (0 <= pf200 && pf200 < 5) return isSand ? (wg ? "SW" : "SP") : (wg ? "GW" : "GP");
  if (5 <= pf200 && pf200 < 12 && isSand) return wg ? (clayFines ? "SW-SC" : "SW-SM") : (clayFines ? "SP-SC" : "SP-SM");
  if (5 <= pf200 && pf200 < 12 && !isSand) return wg ? (clayFines ? "GW-GC" : "GW-GM") : (clayFines ? "GP-GC" : "GP-GM");
  if (12 <= pf200 && pf200 < 50 && fines === "CL-ML") return isSand ? "SM-SC" : "GM-GC";   // never true: Plasticity returns "CL & ML" (kept)
  if (12 <= pf200 && pf200 < 50) return clayFines ? (isSand ? "SC" : "GC") : (isSand ? "SM" : "GM");
  if (pf200 >= 50) return fines;
  throw new XlErr(NA_CODE);
};
const ASTM_TABLE = [["Pt", "Peat"], ["SW", "Well-graded sand"], ["SW", "Well-graded sand with gravel"], ["SP", "Poorly graded sand"], ["SP", "Poorly graded sand with gravel"], ["SW-SM", "Well-graded sand with silt"], ["SW-SM", "Well-graded sand with silt and gravel"], ["SW-SC", "Well-graded sand with clay (or silty clay)"], ["SW-SC", "Well-graded sand with clay and gravel (or silty clay and gravel)"], ["SP-SM", "Poorly graded sand with silt"], ["SP-SM", "Poorly graded sand with silt and gravel"], ["SP-SC", "Poorly graded sand with clay (or silty clay)"], ["SP-SC", "Poorly graded sand with clay and gravel (or silty clay and gravel)"], ["SM", "Silty sand"], ["SM", "Silty sand with gravel"], ["SC", "Clayey sand"], ["SC", "Clayey sand with gravel"], ["SC-SM", "silty, clayey sand"], ["SC-SM", "silty, clayey sand with gravel"], ["GW", "Well-graded gravel"], ["GW", "Well-graded gravel with sand"], ["GP", "Poorly graded gravel"], ["GP", "Poorly graded gravel with sand"], ["GW-GM", "Well-graded gravel with silt"], ["GW-GM", "Well-graded gravel with silt and sand"], ["GW-GC", "Well-graded gravel with clay (or silty clay)"], ["GW-GC", "Well-graded gravel with clay and sand (or silty clay and sand)"], ["GP-GM", "Poorly graded gravel with silt"], ["GP-GM", "Poorly graded gravel with silt and sand"], ["GP-GC", "Poorly graded gravel with clay (or silty clay)"], ["GP-GC", "Poorly graded gravel with clay and sand (or silty clay and sand)"], ["GM", "Silty gravel"], ["GM", "Silty gravel with sand"], ["GC", "Clayey gravel"], ["GC", "Clayey gravel with sand"], ["GC-GM", "silty, clayey gravel"], ["GC-GM", "silty, clayey gravel with sand"], ["CL", "Lean clay"], ["CL", "Lean clay with sand"], ["CL", "Lean clay with gravel"], ["CL", "Sandy lean clay"], ["CL", "Sandy lean clay with gravel"], ["CL", "Gravelly lean clay"], ["CL", "Gravelly lean clay with sand"], ["CL-ML", "Silty clay"], ["CL-ML", "Silty clay with sand"], ["CL-ML", "Silty clay with gravel"], ["CL-ML", "Sandy silty clay"], ["CL-ML", "Sandy silty clay with gravel"], ["CL-ML", "Gravelly silty clay"], ["CL-ML", "Gravelly silty clay with sand"], ["ML", "Silt"], ["ML", "Silt with sand"], ["ML", "Silt with gravel"], ["ML", "Sandy silt"], ["ML", "Sandy silt with gravel"], ["ML", "Gravelly silt"], ["ML", "Gravelly silt with sand"], ["OML", "Organic silt"], ["OML", "Organic silt with sand"], ["OML", "Organic silt with gravel"], ["OML", "Sandy organic silt"], ["OML", "Sandy organic silt with gravel"], ["OML", "Gravelly organic silt"], ["OML", "Gravelly organic silt with sand"], ["OML", "Organic clay"], ["OML", "Organic clay with sand"], ["OML", "Organic clay with gravel"], ["OML", "Sandy organic clay"], ["OML", "Sandy organic clay with gravel"], ["OML", "Gravelly organic clay"], ["OML", "Gravelly organic clay with sand"], ["CH", "Fat clay"], ["CH", "Fat clay with sand"], ["CH", "Fat clay with gravel"], ["CH", "Sandy fat clay"], ["CH", "Sandy fat clay with gravel"], ["CH", "Gravelly fat clay"], ["CH", "Gravelly fat clay with sand"], ["MH", "Elastic silt"], ["MH", "Elastic silt with sand"], ["MH", "Elastic silt with gravel"], ["MH", "Sandy elastic silt"], ["MH", "Sandy elastic silt with gravel"], ["MH", "Gravelly elastic silt"], ["MH", "Gravelly elastic silt with sand"], ["OCH", "Organic clay"], ["OCH", "Organic clay with sand"], ["OCH", "Organic clay with gravel"], ["OCH", "Sandy organic clay"], ["OCH", "Sandy organic clay with gravel"], ["OCH", "Gravelly organic clay"], ["OCH", "Gravelly organic clay with sand"], ["OCH", "Organic silt"], ["OCH", "Organic silt with sand"], ["OCH", "Organic silt with gravel"], ["OCH", "Sandy organic silt"], ["OCH", "Sandy organic silt with gravel"], ["OCH", "Gravelly organic silt"], ["OCH", "Gravelly organic silt with sand"]];
G.astm_table = () => ASTM_TABLE.map((r) => r.slice());
G.classify_ASTM = (list, pf4, pf200, Cu_, Cc_, LL_, PL_, org_) => {
  const o = (v) => (v === undefined ? 0 : num(v));
  const organic = o(org_), Cu = o(Cu_), Cc = o(Cc_), LL = o(LL_), PL = o(PL_);
  const sand = pf4 - pf200, gravel = 100 - pf4;
  const U = G.classify_USC(pf4, pf200, Cu, Cc, LL, PL, organic);
  const USC = organic > 0 && (U === "ML" || U === "MH") ? "O" + U : organic > 0 && U === "MC" ? "O" + U : U;
  const i = pf200 > 85 ? 1 : pf200 > 70 ? (sand >= gravel ? 2 : 3) : sand >= gravel ? (gravel > 15 ? 4 : 5) : (sand > 15 ? 6 : 7);
  const j = pf200 > 50 ? i : sand < 15 ? 1 : 2;
  if (!Array.isArray(list) || !list.length || !Array.isArray(list[0]) || list[0].length < 2) list = ASTM_TABLE;
  const g = list.filter((r) => str(r[0]) === USC);
  if (g.length < j) throw new XlErr(REF_CODE);
  return str(g[j - 1][0]) + " - " + str(g[j - 1][1]);
};
G.classify_ASTM2487 = (pf4, pf200, cu_, cc_, ll_, pl_, ratio_) => {
  const N = (v) => (v === undefined || v === "" ? 0 : (typeof v === "number" ? v : 0));
  const cuV = N(cu_), ccV = N(cc_), llV = N(ll_), plV = N(pl_);
  const rt = ratio_ === undefined || N(ratio_) === 0 ? 1 : N(ratio_);
  const gravel = 100 - pf4, sand = pf4 - pf200, plus = 100 - pf200;
  const piV = llV - plV, above = piV >= 0.73 * (llV - 20);
  const ftype = piV < 4 || !above ? "M" : llV < 50 && piV <= 7 ? "CL-ML" : "C";
  const isG = gravel > sand, gL = isG ? "G" : "S", main = isG ? "gravel" : "sand", other = isG ? "sand" : "gravel";
  const oPct = isG ? sand : gravel, withO = oPct >= 15 ? " with " + other : "", andO = oPct >= 15 ? " and " + other : "";
  const good = cuV >= (isG ? 4 : 6) && ccV >= 1 && ccV <= 3;
  const grd = good ? "W" : "P", grdTxt = good ? "Well-graded " : "Poorly graded ";
  const fm = ftype === "M" ? "M" : "C";
  if (pf200 < 50) {
    if (pf200 < 5) return [gL + grd, grdTxt + main + withO];
    if (pf200 <= 12) return [gL + grd + "-" + gL + fm, grdTxt + main + (fm === "M" ? " with silt" + andO : " with clay" + andO + " (or silty clay" + andO + ")")];
    if (ftype === "M") return [gL + "M", "Silty " + main + withO];
    if (ftype === "C") return [gL + "C", "Clayey " + main + withO];
    return [gL + "C-" + gL + "M", "Silty, clayey " + main + withO];
  }
  const org = rt < 0.75;
  const fSym = org ? (llV < 50 ? "OL" : "OH")
    : llV < 50 ? (!above ? "ML" : piV > 7 ? "CL" : piV >= 4 ? "CL-ML" : "ML") : (above ? "CH" : "MH");
  const base = org ? ((llV < 50 ? piV >= 4 && above : above) ? "organic clay" : "organic silt")
    : ({ CL: "lean clay", "CL-ML": "silty clay", ML: "silt", CH: "fat clay" }[fSym] || "elastic silt");
  const cap = base[0].toUpperCase() + base.slice(1), sg = sand >= gravel;
  const fName = plus < 15 ? cap : plus < 30 ? cap + (sg ? " with sand" : " with gravel")
    : sg ? "Sandy " + base + (gravel >= 15 ? " with gravel" : "") : "Gravelly " + base + (sand >= 15 ? " with sand" : "");
  return [fSym, fName];
};

// ---------------------------------------------------------------- consolidation
G.cutline = (X, Y, slp, icpt) => {
  const rows = [];
  for (let i = 0; i < X.length - 1; i++) {
    const den = Y[i + 1] - Y[i] + slp * (X[i] - X[i + 1]);
    const z = den !== 0 ? (icpt - Y[i] + slp * X[i]) / den : -10;
    if (z > 0 && z <= 1) rows.push([X[i] + z * (X[i + 1] - X[i]), Y[i] + z * (Y[i + 1] - Y[i])]);
  }
  if (!rows.length) throw new XlErr(NA_CODE, "cutline: no intersection");
  return rows;
};
G.consolidation_test = (T, X, ta, tb, hi) => {
  const lbl = ["Consolidation coefficient Cv =", "Initial height =", "Initial reading =", "Reading at consolidation beginning d0 =",
    "Reading at consolidation end d100 =", "Final reading =", "Average drainage length =", "Initial compression ri =",
    "Primary compression rp =", "Secondary compression rs ="];
  const m = T.map((t) => t > ta && t <= tb);
  const tf = filter(T, m), xf = filter(X, m), sq = tf.map(sqrt);
  const a = slope(xf, sq) / 1.15, B = intercept(xf, sq);
  const cut = G.cutline(T.map(sqrt), X, a, B);
  const xc = cut[0][1], tc = xc ** 2;                     // as in the LAMBDA: both from the y of the first cut
  const Cv = 0.848 * xc ** 2 / tc ** 2;
  const xe = X[X.length - 1], xs = X[0], xp = B + (xc - B) / 0.9, dx = xs - xe;
  const vals = [Cv, hi, xs, B, xp, xe, (hi + xe - (B + xp) / 2) / 2, (xs - B) / dx, (B - xp) / dx, (xp - xe) / dx];
  return lbl.map((l, i) => [l, vals[i]]);
};
G.consolidation_Tv = (U) => (U > 0.6 ? -0.085 - 0.933 * log10(1 - U) : PI / 4 * U ** 2);
const CASES = ["uniform", "triangular_top", "triangular_bottom", "vshape", "saw", "sine", "parabolic", "circular"];
function caseIndexOf(c) {
  if (c === undefined) return 1;
  const i = CASES.indexOf(lower(c));
  if (i < 0) throw new XlErr(NA_CODE, "unknown case");
  return i + 1;
}
const termCount = (tvMin) => Math.min(1000, 2 + Math.floor(sqrt(160 / Math.max(tvMin, 1e-6)) / PI));
G.consolidation_U = (tvs, case_) => {
  const ci = caseIndexOf(case_), N = termCount(Math.min(...tvs));
  const cmU = (n) => {
    const np = n * PI, odd = MOD(n, 2) !== 0, sn = odd ? (-1) ** ((n - 1) / 2) : 0;
    switch (ci) {
      case 1: return 8 * sn ** 4 / np ** 2;
      case 2: case 3: return odd ? 8 / np ** 2 : 0;
      case 4: return 16 * sn ** 3 * (np * sn - 2) / np ** 3;
      case 5: return 32 * sn ** 3 / np ** 3;
      case 6: return n === 1 ? 1 : 0;
      case 7: return odd ? 96 / np ** 4 : 0;
      case 8: return odd ? -16 * (-1) ** ((n + 1) / 2) * besselJ1(np / 2) / np ** 2 : 0;
      default: throw new XlErr(NA_CODE);
    }
  };
  const cm = seq(N).map(cmU);
  return tvs.map((tv) => 1 - sum(cm.map((c, k) => { const np = (k + 1) * PI; return c * (np ** 2 * tv / 4 > 700 ? 0 : 1 / exp(np ** 2 * tv / 4)); })));
};
G.consolidation_Ux = (X, tv, case_) => {
  const ci = caseIndexOf(case_), N = termCount(tv);
  const terms = seq(N).map((n) => {
    const np = n * PI, odd = MOD(n, 2) !== 0, sn = odd ? (-1) ** ((n - 1) / 2) : 0;
    let c;
    switch (ci) {
      case 1: c = 4 * sn ** 2 / np; break;
      case 2: c = -2 * (-1) ** n / np; break;
      case 3: c = 2 / np; break;
      case 4: c = 4 * sn * (np * sn - 2) / np ** 2; break;
      case 5: c = 8 * sn / np ** 2; break;
      case 6: c = n === 1 ? 1 : 0; break;
      case 7: c = odd ? -8 * (-1) ** n / np ** 3 : 0; break;
      case 8: c = 2 * sn / n * besselJ1(np / 2); break;
      default: throw new XlErr(NA_CODE);
    }
    const en = np ** 2 * tv / 4 > 700 ? 0 : 1 / exp(np ** 2 * tv / 4);
    return [c * en, np];
  });
  return X.map((x) => sum(terms.map(([ce, np]) => ce * sin(np * x))));
};
G.consolidation_U_approximate = (tv) => { const U = 2 * sqrt(tv / PI); return U < 0.6 ? U : 1 - 8 / PI ** 2 / exp(PI ** 2 / 4 * tv); };
G.consolidation_U_variable = (tv, tc) => (tv < tc ? G.consolidation_U([tv / 2])[0] * tv / tc : G.consolidation_U([tv - tc / 2])[0]);

// ---------------------------------------------------------------- trigonometry in degrees
G.cosd = (x) => cos(rad(x));
G.sind = (x) => sin(rad(x));
G.tand = (x) => tan(rad(x));

// ---------------------------------------------------------------- earth pressure
G.Coulomb_cohesiveforces = (t, phi, alpha, beta, delta, g, h, coh_) => {
  const f = rad(phi), ar = rad(alpha), br = rad(beta), dr = rad(delta);
  const BC = h * sin(br - ar) / sin(br - t) / sin(ar), AB = h * sqrt(1 + (1 / tan(ar)) ** 2);
  const Ka = (1 - sin(f)) / (1 + sin(f)), coh = coh_ === undefined ? 0 : num(coh_);
  const z0 = 2 * coh / g / sqrt(Ka), Cw = coh * Math.max(0, AB - z0), Cs = coh * Math.max(0, BC - z0);
  const W = g * AB * BC * sin(ar - t) / 2, den = sin(f + ar + dr - t);
  const R = (W * sin(ar + dr) - Cw * cos(dr) - Cs * cos(ar + dr - t)) / den;
  const P = (W * sin(t - f) - Cs * cos(f) - Cw * cos(ar + f - t)) / den;
  return [W, R, P, Cs, Cw, z0, P / (g * h ** 2 / 2)];
};
G.Coulomb_forces = (t, phi, alpha, beta, delta, g, h) => {
  const f = rad(phi), ar = rad(alpha), br = rad(beta), dr = rad(delta);
  const BC = h * sin(br - ar) / sin(br - t) / sin(ar), AB = h * sqrt(1 + (1 / tan(ar)) ** 2);
  const W = g * AB * BC * sin(ar - t) / 2, den = sin(t - f - ar - dr);
  const P = W * sin(f - t) / den;
  return [W, -W * sin(ar + dr) / den, P, P / (g * h ** 2 / 2)];
};
G.coulomb_K = (phi, alpha, beta, delta, passive) => {
  const f = rad(phi), ar = rad(alpha), br = rad(beta), dr = rad(delta);
  if (passive === undefined)
    return (sin(ar - f) / sin(ar) / (sqrt(sin(ar + dr)) + sqrt(sin(f + dr) * sin(f - br) / sin(ar - br)))) ** 2;
  return (sin(ar + f) / sin(ar) / (sqrt(sin(ar - dr)) - sqrt(sin(f + dr) * sin(f + br) / sin(ar - br)))) ** 2;
};
G.Coulomb_maximum = (a_, b_, phi, alpha, beta, delta, g, h, iter_) => {
  const niter = 10, gr = (sqrt(5) + 1) / 2;
  let iter = iter_ === undefined ? 1 : num(iter_);
  let A = iter === 1 ? PI / 2 : a_, B = iter === 1 ? rad(beta) + 0.01 : b_;
  while (iter < niter) {                                   // the LAMBDA recursion, written as a loop
    const c = B - (B - A) / gr, d = A + (B - A) / gr;
    const Yc = -G.Coulomb_forces(c, phi, alpha, beta, delta, g, h)[2];
    const yd = -G.Coulomb_forces(d, phi, alpha, beta, delta, g, h)[2];
    if (Yc < yd) B = d; else A = c;
    iter++;
  }
  return [deg((A + B) / 2), ...G.Coulomb_forces((A + B) / 2, phi, alpha, beta, delta, g, h)];
};
G.coulomb_theta = (phi, alpha, beta, delta) => {
  const A = rad(alpha - beta), B = rad(beta - phi), D = rad(alpha + delta + phi - beta);
  const cA = sin(A) * sin(B), cB = sin(A - B), cC = -cos(A) * cos(B), cD = sin(D), cE = -cos(D);
  const qA = cA * cD, qB = 2 * cA * cE, qC = cB * cE - cC * cD;
  const root = (-qB - sqrt(qB ** 2 - 4 * qA * qC)) / (2 * qA);
  return beta + deg(ATAN2(root, 1));
};
G.Rankine_K = (phi, beta) => {
  const f = rad(phi), b = rad(beta), r = sqrt(cos(b) ** 2 - cos(f) ** 2);
  return (cos(b) - r) / (cos(b) + r);
};

// ---------------------------------------------------------------- elastic stresses
G.elastic_circular_radius = (s) => sqrt(1 / (1 - s) ** (2 / 3) - 1);
G.elastic_circular_stress = (r, z, nu_) => {
  const nu = nu_ === undefined ? 0.3 : num(nu_), D = r ** 2 + z ** 2;
  return ["Stress sz, sr =", 1 - 1 / (1 + (r / z) ** 2) ** 1.5, 1 / 2 * (1 + 2 * nu - 2 * (1 + nu) * z / sqrt(D) + z ** 3 / D ** 1.5)];
};
G.elastic_embankment = (b1z, b2z) => (atan(b1z) + (b1z + b2z) / b2z * (atan(b1z + b2z) - atan(b1z))) / PI;
const JANBU0 = {
  1: [[0.1, 0.475, 0.615, 0.892, 1.17, 1.876, 2.965, 4.12, 6.994, 12.393, 21.647, 1000], [1, 0.859, 0.816, 0.755, 0.703, 0.633, 0.586, 0.556, 0.53, 0.509, 0.501, 0.5]],
  2: [[0.1, 0.606, 0.867, 1.154, 1.603, 2.498, 3.836, 6.511, 12.043, 19.585, 30, 1000], [1, 0.869, 0.81, 0.76, 0.706, 0.643, 0.598, 0.562, 0.527, 0.514, 0.508, 0.5]],
  5: [[0.1, 0.972, 1.45, 2.196, 2.645, 4.301, 7.095, 12.393, 21.037, 40, 1000], [1, 0.851, 0.793, 0.729, 0.702, 0.641, 0.588, 0.547, 0.523, 0.51, 0.505]],
  10: [[0.1, 1, 1.389, 2.015, 3.52, 6.149, 10.589, 17.22, 28.406, 40.039, 1000], [1, 0.881, 0.838, 0.789, 0.714, 0.647, 0.593, 0.555, 0.526, 0.516, 0.505]],
  20: [[0.1, 1.65, 2.721, 4.301, 5.891, 10.29, 17.975, 32.774, 67.002, 95.8, 1000], [1, 0.848, 0.799, 0.742, 0.702, 0.643, 0.592, 0.545, 0.517, 0.515, 0.505]],
  50: [[0.1, 0.819, 1.275, 1.931, 2.882, 4.822, 7.301, 9.858, 18.763, 34.703, 67.002, 123.928, 207.382, 400, 1000], [1, 0.929, 0.896, 0.864, 0.831, 0.784, 0.739, 0.705, 0.642, 0.591, 0.551, 0.527, 0.515, 0.51, 0.505]],
  100: [[0.1, 2.228, 3.325, 4.686, 6.895, 10.144, 14.095, 24.621, 40.616, 67.967, 127.523, 213.399, 400, 1000], [1, 0.874, 0.838, 0.809, 0.773, 0.74, 0.706, 0.655, 0.617, 0.581, 0.546, 0.529, 0.51, 0.505]],
  200: [[0.1, 1.137, 2.044, 3.051, 4.24, 5.891, 8.185, 11.872, 16.975, 23.927, 33.246, 48.915, 71.969, 108.96, 177.193, 288.156, 455.392, 660.504], [1, 0.937, 0.894, 0.863, 0.839, 0.812, 0.783, 0.756, 0.727, 0.698, 0.674, 0.642, 0.616, 0.587, 0.56, 0.536, 0.521, 0.51]],
};
const JANBU1 = {
  circle: [[0.1, 0.149, 0.201, 0.284, 0.488, 0.749, 1.288, 2.185, 3.602, 6.023, 12.657, 32.948, 94.786, 319.063, 1000], [0.063, 0.087, 0.128, 0.186, 0.267, 0.349, 0.447, 0.529, 0.57, 0.611, 0.625, 0.63, 0.631, 0.633, 0.64]],
  1: [[0.1, 0.149, 0.201, 0.816, 1.326, 2.249, 3.654, 6.948, 14.6, 1000], [0.063, 0.087, 0.128, 0.395, 0.493, 0.58, 0.644, 0.698, 0.717, 0.739]],
  2: [[0.1, 0.149, 0.201, 0.707, 1.345, 3.079, 6.287, 14.81, 30.678, 85.77, 1000], [0.063, 0.087, 0.128, 0.394, 0.59, 0.787, 0.909, 0.991, 1.028, 1.042, 1.053]],
  5: [[0.1, 0.149, 0.201, 0.284, 0.467, 0.771, 1.345, 2.094, 3.354, 6.469, 14.6, 32.02, 84.554, 210.878, 541.17, 1000], [0.063, 0.087, 0.128, 0.186, 0.32, 0.486, 0.693, 0.884, 1.034, 1.23, 1.358, 1.428, 1.482, 1.496, 1.504, 1.506]],
  10: [[0.1, 0.149, 0.201, 0.284, 0.467, 0.771, 1.345, 2.094, 5.529, 8.984, 18.087, 34.885, 82.173, 199.17, 430.634, 1000], [0.063, 0.087, 0.128, 0.186, 0.32, 0.486, 0.693, 0.884, 1.264, 1.414, 1.559, 1.646, 1.723, 1.765, 1.79, 1.798]],
  20: [[0.1, 0.149, 0.201, 0.284, 0.467, 0.771, 1.345, 2.381, 3.551, 5.222, 8.365, 14.81, 25.847, 43.839, 88.254, 172.667, 368.037], [0.063, 0.087, 0.128, 0.186, 0.32, 0.486, 0.693, 0.947, 1.137, 1.309, 1.534, 1.707, 1.84, 1.95, 2.026, 2.074, 2.099]],
  50: [[0.1, 0.149, 0.201, 0.284, 0.467, 0.771, 1.345, 2.381, 3.551, 5.222, 8.365, 25.48, 42.001, 70.228, 127.931, 268.817, 518.476], [0.063, 0.087, 0.128, 0.186, 0.32, 0.486, 0.693, 0.947, 1.137, 1.309, 1.534, 2.052, 2.213, 2.335, 2.417, 2.481, 2.506]],
  100: [[0.1, 0.149, 0.201, 0.284, 0.467, 0.771, 1.345, 2.381, 3.551, 5.222, 8.365, 89.524, 145.476, 257.545, 469.157, 1000], [0.063, 0.087, 0.128, 0.186, 0.32, 0.486, 0.693, 0.947, 1.137, 1.309, 1.534, 2.633, 2.72, 2.773, 2.798, 2.805]],
};
G.elastic_Janbu_mu0 = (DB, LB) => { const t = JANBU0[num(LB)]; if (!t) throw new XlErr(NA_CODE); return G.interp(DB, t[0], t[1], "logx"); };
G.elastic_Janbu_mu1 = (LB, DB) => {
  const key = typeof DB === "string" && DB.toLowerCase() === "circle" ? "circle" : num(DB);
  const t = JANBU1[key]; if (!t) throw new XlErr(NA_CODE);
  return G.interp(LB, t[0], t[1], "logx");
};
G.elastic_lineload = (x, y) => { const t = 2 / PI / (x ** 2 + y ** 2) ** 2; return ["Stress sx,sy,sxy =", t * x ** 2 * y, t * y ** 3, t * x * y ** 2]; };
G.elastic_lineload_horizontal = (x, y) => { const t = 2 / PI / (x ** 2 + y ** 2) ** 2; return ["Stress sx,sy,sxy =", t * x ** 3, t * x * y ** 2, t * x ** 2 * y]; };
G.elastic_Newmark_factor = (m, n) => { const v = 1 + m ** 2 + n ** 2; return 1 / (2 * PI) * (m * n * (1 + v) / sqrt(v) / (1 + m ** 2) / (1 + n ** 2) + atan(m * n / sqrt(v))); };
G.elastic_Newmark_stress = (x_, y_, z_, pt) => {
  const z = z_ === 0 ? 1 : z_, x = x_ / z, y = y_ / z;
  const xa = pt[0][0] / z, xb = pt[2][0] / z, ya = pt[0][1] / z, yb = pt[2][1] / z;
  const F = G.elastic_Newmark_factor;
  return SIGN((x - xa) * (y - ya)) * F(abs(x - xa), abs(y - ya)) + SIGN((x - xb) * (y - yb)) * F(abs(x - xb), abs(y - yb))
    + SIGN((xa - x) * (y - yb)) * F(abs(xa - x), abs(y - yb)) + SIGN((xb - x) * (y - ya)) * F(abs(xb - x), abs(y - ya));
};
G.elastic_pointload = (r, z, nu_) => {
  const nu = nu_ === undefined ? 0.3 : num(nu_), D = r ** 2 + z ** 2;
  return ["Stress sz,sr,st,srz =", 3 / (2 * PI * z ** 2) / (1 + (r / z) ** 2) ** 2.5,
    PI / 2 * (3 * r ** 2 * z / D ** 2.5 - (1 - 2 * nu) / (D + z * sqrt(D))),
    -1 / 2 / PI * (1 - 2 * nu) * (z / D ** 1.5 - 1 / (r ** 2 + z ** 2 + z * sqrt(D))),
    3 / 2 / PI * r * z ** 2 / D ** 2.5];
};
G.elastic_sector_rowValue = (t, poly) => {          // one ray of elastic_sector
  let rs;
  try { rs = G.intersection_polygon_ray(poly, [0, 0], t); } catch (e) { return 0; }
  const s = rs.map((r) => 1 - 1 / (1 + r ** 2) ** 1.5);
  if (s.length === 2) return abs(s[1] - s[0]);
  if (s.length === 1) return s[0];
  throw new XlErr(VALUE_CODE, "ray crosses the polygon more than twice");
};
G.elastic_sector = (x, y, z, poly) => {
  const n = 360, rel = poly.map(([px, py]) => [(px - x) / z, (py - y) / z]);
  let s = 0;
  for (let k = 0; k <= n; k++) s += G.elastic_sector_rowValue(rad(k), rel);   // 361 rays, as in the LAMBDA
  return s / n;
};
G.elastic_sector_slice = (ra, rb, z, dt) => dt * (1 / (1 + ra ** 2 / z ** 2) ** 1.5 - 1 / (1 + rb ** 2 / z ** 2) ** 1.5);
const stripCore = (B, x, y) => {
  const tm = ATAN2(y, x - B / 2), tp = ATAN2(y, x + B / 2);
  return [(tp - tm + sin(tp - tm) * cos(tp + tm)) / PI, (tp - tm - sin(tp - tm) * cos(tp + tm)) / PI, sin(tp - tm) * sin(tp + tm) / PI];
};
G.elastic_stripload = (B, x, y) => { const [sy, sx, sxy] = stripCore(B, x, y); return ["Stress sy,sx and sxy", sy, sx, sxy]; };
G.elastic_stripload_xy = (B, x, y) => { const [sy, sx, sxy] = stripCore(B, x, y); return ["sigx, sigy and tauxy =", sx, sy, sxy]; };
G.elastic_triangle = (B, x, y) => {
  const a = atan(x / y) + atan((B - x) / y), r = x ** 2 + y ** 2, D = r + B ** 2 - 2 * B * x;
  return ["Stress sy, sx, sxy =", (x * a + y * (1 - (r - B * x) / D)) / B / PI,
    y * (ln(D / r) + r / D - 1 + x / y * a - B * x / D) / B / PI, (y ** 2 / D - y / B * a) / PI];
};

// ---------------------------------------------------------------- interpolation
/** Interpolate one value (Xa, Ya numeric vectors, any order). */
G.interp = (X, Xa, Ya, mode, outside) => {
  const md = mode === undefined ? "LINEAR" : upper(mode), od = outside === undefined ? "ERROR" : upper(outside);
  const lx = md === "LOGX" || md === "LOGXY", ly = md === "LOGY" || md === "LOGXY";
  const idx = Xa.map((_, i) => i).sort((p, q) => Xa[p] - Xa[q]);           // SORTBY (stable)
  const xs = idx.map((i) => Xa[i]), ys = idx.map((i) => Ya[i]), n = xs.length;
  const xlo = xs[0], xhi = xs[n - 1];
  const q = od === "CLAMP" ? (X < xlo ? xlo : X > xhi ? xhi : X) : X;
  let iRaw = 0;                                                            // XMATCH(q, xs, -1, 2)
  for (let k = 0; k < n; k++) if (xs[k] <= q) iRaw = k + 1;
  if (iRaw === 0) iRaw = 1;
  const i = iRaw > n - 1 ? n - 1 : iRaw;
  const xA = xs[i - 1], xB = xs[i], yA = ys[i - 1], yB = ys[i];
  if ((q < xlo || q > xhi) && od === "ERROR") return NAV;
  if (xB === xA) return yA;
  const t = lx ? ln(q / xA) / ln(xB / xA) : (q - xA) / (xB - xA);
  return ly ? yA * (yB / yA) ** t : yA + t * (yB - yA);
};
/** interp over a vector of query values. */
G.interpV = (Xs, Xa, Ya, mode, outside) => Xs.map((x) => G.interp(x, Xa, Ya, mode, outside));
G.interpolate_table = (x, Y, XY) => {
  const Xa = XY[0].slice(1).map(num), Ya = XY.slice(1).map((r) => num(r[0])), Za = XY.slice(1).map((r) => r.slice(1).map(num));
  const c = xmatchNextSmaller(x, Xa), r = xmatchNextSmaller(Y, Ya);
  if (!c || !r) throw new XlErr(NA_CODE);
  const Xl = Xa[c - 1], Xu = Xa[c], Yl = Ya[r - 1], Yu = Ya[r];
  const a = (Xu - Xl) * (Yu - Yl);
  const Nll = (Xu - x) * (Yu - Y), Nul = (x - Xl) * (Yu - Y), Nuu = (x - Xl) * (Y - Yl), Nlu = (Xu - x) * (Y - Yl);
  const Zll = Za[r - 1][c - 1], Zlu = Za[r][c - 1], Zuu = Za[r][c];
  const Zul = Za[r][c - 1];   // NOTE: kept from the LAMBDA (INDEX(Za, r+1, c)); probably meant Za[r-1][c] — see README
  return (Nll * Zll + Nlu * Zlu + Nuu * Zuu + Nul * Zul) / a;
};

// ---------------------------------------------------------------- geometry
G.intersection_polygon_line = (poly, pt, t) => {
  const n = poly.length, closed = poly[0][0] === poly[n - 1][0] && poly[0][1] === poly[n - 1][1];
  const P = closed ? poly : [...poly, poly[0]];
  const [xp, yp] = pt, bs = [];
  for (let i = 0; i < P.length - 1; i++) {
    const [Xa, Ya] = P[i], [Xb, Yb] = P[i + 1], dX = Xb - Xa, dY = Yb - Ya;
    const det = sin(t) * dX - cos(t) * dY;
    const a = det !== 0 ? (sin(t) * (Xb - xp) + cos(t) * (yp - Yb)) / det : 0;
    const b = det !== 0 ? (a >= 0 && a <= 1 ? 1 : 0) * (-Xa * Yb + Ya * Xb + xp * dY - yp * dX) / det : 0;
    bs.push(b);
  }
  const bu = [...new Set(bs.filter((b) => b !== 0))];
  if (!bu.length) return [[0, 0]];
  return bu.map((b) => [xp + b * cos(t), yp + b * sin(t)]);
};
G.intersection_polygon_ray = (poly, pt, t) => {
  const XY = G.intersection_polygon_line(poly, pt, t);
  const r = XY.map(([x, y]) => { const X = x - pt[0], Y = y - pt[1]; return ROUND(sqrt(X ** 2 + Y ** 2), 8) * (X * cos(t) + Y * sin(t) > 0 ? 1 : 0); });
  const u = sortUnique(r.filter((v) => v !== 0));
  if (!u.length) throw new XlErr(NA_CODE, "ray does not cross the polygon");
  return u;
};
G.intersection_polygon_line_slope = (poly, M, t) => {
  const P = [...uniqueRows(poly), poly[0]], Nv = [cos(t), sin(t)], bs = [];
  for (let i = 0; i < P.length - 1; i++) {
    const AB = [P[i + 1][0] - P[i][0], P[i + 1][1] - P[i][1]], AM = [M[0] - P[i][0], M[1] - P[i][1]];
    const det = Nv[0] * AB[1] - Nv[1] * AB[0];
    const a = det !== 0 ? (Nv[0] * AM[1] - Nv[1] * AM[0]) / det : 0;
    bs.push(det !== 0 ? (a >= 0 && a <= 1 ? 1 : 0) * (AB[0] * AM[1] - AB[1] * AM[0]) / det : 0);
  }
  const bu = [...new Set(bs.filter((b) => b !== 0))];
  if (!bu.length) return [[0, 0]];
  return bu.map((b) => [M[0] + b * Nv[0], M[1] + b * Nv[1]]);
};
G.polygon_area_centroid = (poly) => {
  const n = poly.length, closed = poly[0][0] === poly[n - 1][0] && poly[0][1] === poly[n - 1][1];
  const P = closed ? poly : [...poly, poly[0]];
  let A2 = 0, sx = 0, sy = 0;
  for (let i = 0; i < P.length - 1; i++) {
    const [Xb, Yb] = P[i], [Xt, Yt] = P[i + 1], Z = Xt * Yb - Xb * Yt;
    A2 += Z; sx += (Xt + Xb) * Z; sy += (Yt + Yb) * Z;
  }
  const a = A2 / 2;
  return [["Centroid = ", sx / 6 / a, sy / 6 / a], ["Area = ", abs(a), ""]];
};
G.polygon_weightMoment = (XY, g, Pt) => {
  const CA = G.polygon_area_centroid(XY), W = g * CA[1][1];
  return [W, W * (Pt[0] - CA[0][1])];
};
G.Polygon_inside = (x, y, xa, ya) => {
  let count = 0;
  for (let i = 0; i < xa.length - 1; i++) {
    const a = (x - xa[i]) / (xa[i + 1] - xa[i]), b = ya[i] + a * (ya[i + 1] - ya[i]);
    if (a >= 0 && a <= 1 && b >= y) count++;
  }
  return count === 1;
};
G.polyline_subdivide = (X, pl) => {
  const Xa = pl.map((r) => r[0]), Ya = pl.map((r) => r[1]);
  const i = xmatchNextSmaller(X, Xa);
  if (!i || i >= Xa.length) throw new XlErr(NA_CODE);
  const S = (Ya[i] - Ya[i - 1]) / (Xa[i] - Xa[i - 1]);
  return Ya[i - 1] + S * (X - Xa[i - 1]);
};
G.distance_point_polyline = (M, pl) => pl.slice(0, -1).map((A, i) => {
  const B = pl[i + 1], AB = (B[0] - A[0]) ** 2 + (B[1] - A[1]) ** 2, AM = (A[0] - M[0]) ** 2 + (A[1] - M[1]) ** 2;
  const ABAM = (A[0] - M[0]) * (B[0] - A[0]) + (A[1] - M[1]) * (B[1] - A[1]);
  return sqrt(AM - ABAM ** 2 / AB);
});
G.intersection_polyline_circle = (pline, center, radius) => {
  const pl = pline.map(([x, y]) => [x - center[0], y - center[1]]), out = [], second = [];
  for (let i = 0; i < pl.length - 1; i++) {
    const A = pl[i], AB = [pl[i + 1][0] - A[0], pl[i + 1][1] - A[1]];
    const OAs = A[0] ** 2 + A[1] ** 2, ABs = AB[0] ** 2 + AB[1] ** 2, AAB = A[0] * AB[0] + A[1] * AB[1];
    const dis = AAB ** 2 - ABs * (OAs - radius ** 2);
    const ap = dis >= 0 ? (-AAB - sqrt(dis)) / ABs : -1, bp = dis >= 0 ? (-AAB + sqrt(dis)) / ABs : -1;
    out.push([A, AB, ap]); second.push([A, AB, bp]);
  }
  const rows = [...out, ...second].filter(([, , s]) => s >= 0 && s <= 1);   // VSTACK(first roots, second roots) then FILTER
  if (!rows.length) throw new XlErr(NA_CODE, "no intersection");
  return rows.map(([A, AB, s]) => [A[0] + s * AB[0] + center[0], A[1] + s * AB[1] + center[1]]);
};
/** Parameter (along segment C-D) of the first crossing with polyline ql; 0 if none. */
G.segment_param = (C, D, ql) => {
  const CD = [D[0] - C[0], D[1] - C[1]];
  for (let i = 0; i < ql.length - 1; i++) {
    const A = ql[i], AB = [ql[i + 1][0] - A[0], ql[i + 1][1] - A[1]], AC = [C[0] - A[0], C[1] - A[1]];
    const det = CD[0] * AB[1] - CD[1] * AB[0];
    const a = det !== 0 ? (CD[0] * AC[1] - CD[1] * AC[0]) / det : 0;
    const b = det !== 0 ? (a >= 0 && a <= 1 ? 1 : 0) * (AB[0] * AC[1] - AB[1] * AC[0]) / det : 0;
    if (a >= 0 && a <= 1 && b >= 0 && b <= 1) return b;
  }
  return 0;
};
G.intersection_segment_polyline = (rowVals) => {
  const P = [];
  for (let i = 0; i + 1 < rowVals.length; i += 2) P.push([rowVals[i], rowVals[i + 1]]);
  const ql = P.slice(2), C = P[0], D = P[1], out = [];
  for (let i = 0; i < ql.length - 1; i++) {                  // full column of b values, as in the LAMBDA
    const A = ql[i], AB = [ql[i + 1][0] - A[0], ql[i + 1][1] - A[1]], AC = [C[0] - A[0], C[1] - A[1]];
    const CD = [D[0] - C[0], D[1] - C[1]], det = CD[0] * AB[1] - CD[1] * AB[0];
    const a = det !== 0 ? (CD[0] * AC[1] - CD[1] * AC[0]) / det : 0;
    const b = det !== 0 ? (a >= 0 && a <= 1 ? 1 : 0) * (AB[0] * AC[1] - AB[1] * AC[0]) / det : 0;
    if (a >= 0 && a <= 1 && b >= 0 && b <= 1) out.push(b);
  }
  return out.length ? out : [0];
};
G.intersection_polyline_polyline = (pl, ql, m) => {
  const mm = m === undefined ? 0 : num(m), rows = [];
  for (let i = 0; i < pl.length - 1; i++) {
    const a = G.segment_param(pl[i], pl[i + 1], ql);
    if (a !== 0) rows.push([a, pl[i][0] + a * (pl[i + 1][0] - pl[i][0]), pl[i][1] + a * (pl[i + 1][1] - pl[i][1])]);
  }
  if (!rows.length) throw new XlErr(NA_CODE, "no intersection");
  return mm > 0 ? sortRows(rows, 1, true).slice(0, mm) : rows;
};
G.intersection_spiral_segment = (AB_, pt, r, phi, x_, iter_) => {
  const niter = 10, ab = AB_.map(([x, y]) => [x - pt[0], y - pt[1]]);
  let iter = iter_ === undefined ? 1 : num(iter_), X = x_ === undefined ? ab[0][0] : num(x_);
  const d = tan(rad(phi)), S = (ab[1][1] - ab[0][1]) / (ab[1][0] - ab[0][0]);
  for (;;) {                                                // Newton iterations (the LAMBDA recursion)
    const Y = ab[0][1] + S * (X - ab[0][0]);
    const f = ln(X ** 2 + Y ** 2) - 2 * ln(r) + 2 * d * atan(Y / X);
    const df = 2 * (X * (1 + d * S) + Y * (S - d)) / (X ** 2 + Y ** 2);
    const xn = X - f / df, yn = ab[0][1] + S * (xn - ab[0][0]);
    if (iter > niter) return [pt[0] + xn, pt[1] + yn];
    X = xn; iter++;
  }
};
G.circle = (pt, r, ta_, tb_, n_) => {
  const n = n_ === undefined ? 100 : num(n_), ta = ta_ === undefined ? 0 : rad(ta_), tb = tb_ === undefined ? 2 * PI : rad(tb_);
  return [["x", "y"], ...seq(n + 1, 0).map((k) => { const a = ta + k * (tb - ta) / n; return [pt[0] + r * cos(a), pt[1] + r * sin(a)]; })];
};
G.logSpiral = (pt, r, phi, ta_, tb_, n_) => {
  const d = tan(rad(phi)), ta = ta_ === undefined ? 0 : rad(ta_), tb = tb_ === undefined ? -PI : rad(tb_), n = n_ === undefined ? 100 : num(n_);
  return seq(n + 1, 0).map((k) => { const t = ta + k * (tb - ta) / n; return [pt[0] + r * exp(-d * t) * cos(t), pt[1] + r * exp(-d * t) * sin(t)]; });
};
G.minimum = (A, B, iter_) => {
  const gr = (sqrt(5) + 1) / 2; let iter = iter_ === undefined ? 1 : num(iter_);
  while (iter < 10) {
    const c = B - (B - A) / gr, d = A + (B - A) / gr;
    if (c ** 3 + 5 * c ** 2 - 10 * c < d ** 3 + 5 * d ** 2 - 10 * d) B = d; else A = c;
    iter++;
  }
  return (A + B) / 2;
};

// ---------------------------------------------------------------- stresses and Mohr circle
G.Stress_Mohr = (sa, ta, alpha, sm) => {
  const ar = rad(alpha), am = 2 * (sm - sa) * cos(ar) - 2 * ta * sin(ar);
  return [sa + am * cos(ar), ta + am * sin(ar)];
};
const MOHR_HDR = ["", "Normal stress", "Shear stress", "Surface orientation (deg)"];
G.stress_Mohr_failure = (sa, ta, A, phi, coh, passive) => {
  const f = rad(phi), sgn = passive === undefined ? -1 : 1;
  const sm = sa / cos(f) ** 2 + coh * tan(f) + sgn * sqrt((coh + sa * tan(f)) ** 2 - ta ** 2) / cos(f);
  const R = sqrt((sa - sm) ** 2 + ta ** 2), ar = A === 0 ? 0.000001 : rad(A);
  const am = 2 * (sm - sa) * cos(ar) - 2 * ta * sin(ar), sp = sa + am * cos(ar), tp = ta + am * sin(ar);
  const s1 = sm + R, s3 = sm - R;
  const sf = sm * cos(f) ** 2 - coh * sin(2 * f) / 2, tfp = coh * cos(f) ** 2 + sm * sin(f) * cos(f);
  return [["Center =", sm, "", ""], ["Radius =", R, "", ""], MOHR_HDR,
    ["Pole", sp, tp, deg(ATAN2(sm - sp, -tp)) + 90],
    ["Major principal stress", s1, 0, deg(ATAN2(s1 - sp, -tp))],
    ["Minor principal stress", s3, 0, deg(ATAN2(s3 - sp, -tp))],
    ["Minor principal stress", s3, 0, deg(ATAN2(s3 - sp, -tp))],       // repeated row kept from the LAMBDA
    ["Failure stress", sf, tfp, deg(ATAN2(sf - sp, tfp - tp))],
    ["Failure stress", sf, -tfp, deg(ATAN2(sf - sp, -tfp - tp))]];
};
G.stress_Mohr_pole = (sa, ta, A, sb) => {
  const tb = -ta, B = A + 90, sm = (sa + sb) / 2, R = sqrt((sa - sb) ** 2 / 4 + ta ** 2), ar = A === 0 ? 0.000001 : rad(A);
  const am = 2 * (sm - sa) * cos(ar) - 2 * ta * sin(ar), sp = sa + am * cos(ar), tp = ta + am * sin(ar);
  const s1 = sm + R, s3 = sm - R;
  return [["Center =", sm, "", ""], ["Radius =", R, "", ""], MOHR_HDR, ["Stress b", sb, tb, B],
    ["Pole", sp, tp, deg(ATAN2(sm - sp, -tp)) + 90],
    ["Major principal stress", s1, 0, deg(ATAN2(s1 - sp, -tp))],
    ["Minor principal stress", s3, 0, deg(ATAN2(s3 - sp, -tp))]];
};
G.stress_Mohrcircle = (sm, r) => [["sig", "tau"], ...seq(21).map((k) => { const a = k * PI / 10; return [sm + r * cos(a), r * sin(a)]; })];
G.Rankine_stress = (phi, coh, alpha, beta, gz, passive) => {
  const br = rad(beta), sa = gz * cos(br) ** 2, ta = gz * sin(br) * cos(br);
  const smf = G.stress_Mohr_failure(sa, ta, beta, phi, coh, passive === undefined ? undefined : 1);
  const sm = smf[0][1], sp = smf[3][1], tp = smf[3][2];
  const [sb, tb] = G.Stress_Mohr(sp, tp, alpha, sm);
  return [["Lateral pressure coefficient and inclination =", sb / sa, deg(ATAN2(sb, tb))],
    ["Normal and shear stress parallel to slope =", sa, ta], ["Normal and shear stress at pole =", sp, tp],
    ["Normal and shear stress on wall =", sb, tb]];
};
G.stress_ForceMoment = (A, B, sa, sb) => {
  const v = [B[0] - A[0], B[1] - A[1]], AB = sqrt(v[0] ** 2 + v[1] ** 2), sm = [(sa[0] + sb[0]) / 2, (sa[1] + sb[1]) / 2];
  const fx = -(sm[0] * v[1] - sm[1] * v[0]), fy = sm[0] * v[0] + sm[1] * v[1];
  const moment = (sa[0] + 2 * sb[0]) * AB ** 2 / 6, zP = (sa[0] + 2 * sb[0]) / sm[0] / 6;
  return [["Force =", fx, fy], ["Moment =", moment, ""], ["Point of application =", A[0] * (1 - zP) + zP * B[0], A[1] * (1 - zP) + zP * B[1]]];
};
G.vector_line = (ptA, ptB, sigA, sigB, M, s) => {
  const AB = [ptB[0] - ptA[0], ptB[1] - ptA[1]], l = sqrt(AB[0] ** 2 + AB[1] ** 2), t = [AB[0] / l, AB[1] / l], nn = [-t[1], t[0]];
  const vA = [nn[0] * sigA[0] + t[0] * sigA[1], nn[1] * sigA[0] + t[1] * sigA[1]];
  const vB = [nn[0] * sigB[0] + t[0] * sigB[1], nn[1] * sigB[0] + t[1] * sigB[1]];
  const rows = [];
  for (let k = 0; k <= M; k++) {
    const a = k / M, p = [(1 - a) * ptA[0] + a * ptB[0], (1 - a) * ptA[1] + a * ptB[1]];
    const v = [(1 - a) * vA[0] + a * vB[0], (1 - a) * vA[1] + a * vB[1]];
    rows.push(p, [p[0] + s * v[0], p[1] + s * v[1]], [NAV, NAV]);
  }
  return rows;
};
G.water_pressure = (depth, wt, unitWeight) => { const gw = unitWeight === undefined ? 9.81 : num(unitWeight); return depth >= wt ? gw * (depth - wt) : 0; };
G.pressure_forceMoment = (d, p, d0) => {
  let F = 0, M = 0;
  for (let i = 0; i < d.length - 1; i++) {
    const db = d[i], dt = d[i + 1], pb = p[i], pt = p[i + 1];
    F += (pt + pb) * (dt - db) / 2;
    M += (dt - db) * (db * (2 * pb + pt) + dt * (2 * pt + pb) - 3 * d0 * (pt + pb)) / 6;
  }
  return [F, M];
};

// ---------------------------------------------------------------- SPT
G.SPT_bearing = (z, N, B, D, author) => {
  const Na = N.length === 1 ? N[0] : averageBetween(N, z, D, B + D);
  const a = xmatchNextLarger(lower(author), ["bowles", "meyerhof", "terzaghi"]);
  const k = 1 + 0.33 * D / B, kB = ((B + 0.3) / B) ** 2, q6 = 10.417 * Na - 20.605, q0 = 13.032 * Na - 16.488, B0 = 1.24;
  if (a === 1) return k * Na * (12.5 * kB >= 20 ? 20 : 12.5 * kB);
  if (a === 2) return k * Na * (8.94 * kB >= 13.41 ? 13.41 : 8.94 * kB);
  if (a === 3) return B < B0 ? q0 : B > 6 ? q6 : q6 + (q0 - q6) * ((B - B0) / (6 - B0) - 1) ** 2;
  return 0;
};
G.SPT_BurlandBurbidge_settlement = (z, N, q, B, LB, D_, H, sc) => {
  const D = D_ === undefined ? 0 : num(D_), ZI = 0.9 * B ** 0.79;
  const Na = averageBetween(N, z, D, ZI + D), Ic = 1.71 / Na ** 1.4;
  const fs = LB === undefined ? 1 : (1.25 * num(LB) / (num(LB) + 0.25)) ** 2;
  const fl = H === undefined ? 1 : num(H) / ZI * (2 - num(H) / ZI);
  const qc = sc === undefined ? q : q > num(sc) ? q - 2 * num(sc) / 3 : q;
  return fs * fl * B ** 0.7 * Ic * qc;
};
G.SPT_correction_depth = (s) => { const CN = s === 0 ? 2 : 10 / sqrt(s); return CN < 2 ? CN : 2; };
G.SPT_Correlation_friction = (N, author, s, pa) => {
  const a = xmatchNextLarger(lower(author), ["peck", "schmertman", "hatanaka"]);
  const noS = s === undefined, noSpa = noS || pa === undefined;
  const CN = G.SPT_correction_depth(noS ? 0 : num(s));
  if (a === 1) return 27.1 + 0.3 * N - 0.00054 * N ** 2;
  if (a === 2) return noSpa ? NAV : deg(atan((N / (12.2 + 20.3 * num(s) / num(pa))) ** 0.34));
  if (a === 3) return noS ? NAV : sqrt(20 * N * CN) + 20;
  return 0;
};
G.SPT_frictionAngle = (N, spa_, author_) => {
  const author = author_ === undefined ? "peck" : lower(author_), spa = spa_ === undefined ? 1 : num(spa_);
  if (author === "peck") return 27.1 + 0.3 * N - 0.00054 * N ** 2;
  if (author === "schmertman") return deg(atan((N / (12.2 + 20.3 * spa)) ** 0.34));
  if (author === "hatanaka") return sqrt(20 * N) + 20;
  return NAV;
};
G.SPT_bearing_direct = (N, B, D, author) => {
  const kd = B === 0 ? 1 : Math.min(1 + 0.33 * D / B, 1.33), q6 = 10.417 * N - 20.605, q0 = 13.032 * N - 16.488, B0 = 1.24;
  const a = lower(author);
  if (a === "bowles") return 18.7 * N * kd * (1 + 0.305 / Math.max(B, 1.22)) ** 2;
  if (a === "meyerhof") return 12.5 * N * kd * (1 + 0.305 / Math.max(B, 1.22)) ** 2;
  if (a === "terzaghi") return B < B0 ? q0 : B > 6 ? q6 : q6 + (q0 - q6) * ((B - B0) / (6 - B0) - 1) ** 2;
  return NAV;
};
G.stepProfile = (x, y) => {
  const n = x.length, xs = [x[0], ...x.slice(0, -1)], xa = x.map((v, i) => (v + xs[i]) / 2);
  const xd = xa.flatMap((v) => [v, v]), xt = [...xd.slice(1), x[n - 1]], yd = y.flatMap((v) => [v, v]);
  return xt.map((v, i) => [v, yd[i]]);
};
G.Step_table = (x, y) => {
  const n = x.length, rows = [];
  for (let k = 1; k <= 2 * n; k++) {
    const ix = Math.floor((k + 1) / 2), iy = Math.floor(k / 2);    // INT((k+1)/2), INT(((k-1)+1)/2)
    rows.push([x[ix - 1], iy === 0 ? y[0] : y[iy - 1]]);          // first y: INDEX(y, 0) in the LAMBDA; y1 used here
  }
  return rows;
};

// ---------------------------------------------------------------- CPT
const CPT_SOILS = ["clay", "silt", "sand", "gravel", "chalk"], CPT_PILES = ["drilled", "cast", "steel", "concrete"];
const CPT_BETA = [[0.000001, 0.000001, 0.000001, 200, 200, 200, 125, 80], [0.000001, 100, 0.000001, 250, 250, 300, 125, 100], [0.000001, 120, 150, 300, 300, 300, 1, 1], [0.000001, 75, 0.000001, 150, 150, 150, 1, 1]];
const CPT_QMAX = [[15, 40, 40, 1000000, 1000000, 120, 40, 120], [15, 40, 40, 1000000, 40, 120, 40, 80], [15, 40, 40, 1000000, 1000000, 120, 1, 1], [15, 80, 80, 1000000, 1000000, 120, 1, 1]];
const CPT_KC = [[0.4, 0.4, 0.4, 0.15, 0.15, 0.15, 0.2, 0.3], [0.55, 0.55, 0.55, 0.5, 0.5, 0.5, 0.3, 0.45]];
G.classify_CPT_LCPC = (qc, soil, pile) => {
  const s = xmatchNextLarger(lower(soil), CPT_SOILS), k = xmatchNextLarger(lower(pile), CPT_PILES);
  if (!k) throw new XlErr(VALUE_CODE, "unknown pile type");
  const j = s <= 2 ? (qc < 3 ? 1 : qc > 6 ? 3 : 2) : s === 5 ? (qc < 5 ? 7 : 8) : (qc < 5 ? 4 : qc > 20 ? 6 : 5);
  return [["A", "B", "C", "A", "B", "C", "A", "B"][j - 1], CPT_KC[k <= 2 ? 0 : 1][j - 1], CPT_BETA[k - 1][j - 1], CPT_QMAX[k - 1][j - 1] / 1000];
};
const PMT_SOILS = ["clay", "silt", "sand", "gravel", "chalk", "marl", "rock"];
const PMT_PILES = ["dry", "mud", "temporary", "permanent", "pier", "steel", "concrete", "moulded", "coated", "low", "high"];
const PMT_KP = [[1.1, 1.2, 1.3, 1, 1.1, 1.2, 1.1, 1.4, 1.8, 0, 1.8, 1.5, 0], [1.4, 1.5, 1.6, 4.2, 3.7, 3.2, 1.6, 2.2, 2.6, 0, 2.6, 2, 0]];
const PMT_QI = [[1, 1, 2, 0, 0, 0, 1, 3, 4, 3, 4, 6, 6], [1, 1, 1, 1, 2, 3, 1, 3, 4, 3, 4, 6, 6], [1, 1, 1, 1, 2, 3, 1, 2, 3, 3, 4, 0, 0], [1, 1, 1, 1, 1, 2, 0, 0, 0, 2, 3, 0, 0],
  [1, 2, 3, 0, 0, 0, 1, 2, 3, 4, 5, 6, 6], [1, 2, 2, 2, 2, 3, 0, 0, 0, 3, 4, 4, 4], [1, 2, 2, 3, 3, 3, 0, 0, 0, 3, 4, 4, 4], [1, 2, 2, 2, 2, 3, 1, 2, 3, 3, 4, 0, 0],
  [1, 2, 2, 3, 3, 3, 0, 0, 0, 3, 4, 0, 0], [1, 2, 2, 3, 3, 3, 2, 3, 4, 5, 5, 0, 0], [0, 4, 5, 5, 5, 6, 0, 5, 6, 6, 6, 7, 7]];
G.classify_PMT_LCPC = (pl, soil, pile) => {
  const s = xmatchNextLarger(lower(soil), PMT_SOILS), k = xmatchNextLarger(lower(pile), PMT_PILES);
  if (!k) throw new XlErr(VALUE_CODE, "unknown pile type");
  const j = s <= 2 ? (pl < 0.7 ? 1 : pl > 2.5 ? 3 : 2) : s <= 4 ? (pl < 0.5 ? 4 : pl > 2.5 ? 6 : 5)
    : s === 5 ? (pl < 0.7 ? 7 : pl > 3 ? 9 : 8) : s === 6 ? (pl > 4.5 ? 11 : 10) : (pl > 4.5 ? 13 : 12);
  return [PMT_KP[k <= 5 ? 0 : 1][j - 1], PMT_QI[k - 1][j - 1]];
};
G.CPT_average_qc = (z, qc, D, B, author) => {
  const a = xmatchNextLarger(lower(author), ["meyerhof", "robertson"]), Db = D + (a === 1 ? 1.5 : 1) * B;
  return average(filter(qc, z.map((v) => (v - D) * (v - Db) <= 0), [0]));
};
G.CPT_bearing_Meyerhof_k = (B, D, dw, m) => {
  const x = (dw - D) / B / 3, C1 = 0.5 + (x > 0.5 ? 0.5 : x), C2 = 0.5 + (dw / D / 2 > 0.5 ? 0.5 : dw / D / 2);
  return B / (m === undefined ? 40 : 12.19) * (C1 + C2 * D / B);
};
G.CPT_bearing_Robertson_k = (B, D, shape, density) => {
  const r = xmatchNextLarger(lower(density), ["dense", "loose"]), s = xmatchNextLarger(lower(shape), ["square", "circular"]);
  const x = B / D > 0 ? log10(B / D) : 0;
  if (r === 1 && s === 1) return -0.246 * x ** 5 + 0.3589 * x ** 4 - 0.0096 * x ** 3 + 0.0482 * x ** 2 - 0.1284 * x + 0.2738;
  if (r === 1 && s === 2) return 0.143 * x ** 4 - 0.1806 * x ** 3 + 0.1497 * x ** 2 - 0.119 * x + 0.2466;
  if (r === 2 && s === 1) return 0.0352 * x ** 4 - 0.0868 * x ** 3 + 0.1904 * x ** 2 - 0.1779 * x + 0.2325;
  if (r === 2 && s === 2) return 0.00253 * x ** 5 + 0.0351 * x ** 4 - 0.1409 * x ** 3 + 0.1905 * x ** 2 - 0.1684 * x + 0.2172;
  return 0;
};
G.CPT_friction_angle = (z, qc, g, author) => {
  const a = xmatchNextLarger(lower(author), ["robertson", "kulhawy", "ricceri", "lee"]), q = qc / (g * z);
  return a === 1 ? deg(atan((log10(q) + 0.29) / 2.68)) : a === 2 ? deg(atan(0.1 + 0.38 * log10(q)))
    : a === 3 ? deg(atan(0.38 + 0.27 * log10(q))) : a === 4 ? 15.575 * q ** 0.1714 : 0;
};
/** Layer table from a value column (shared by the Olsen and Robertson layer functions). */
function layersOf(d, v) {
  const rows = [];
  for (let i = 0; i < d.length; i++) if (v[i] !== (i === 0 ? 0 : v[i - 1])) rows.push([d[i], v[i]]);
  rows.push([d[d.length - 1], v[v.length - 1]]);
  return rows.map(([dd, vv], i) => [dd - (i === 0 ? 0 : rows[i - 1][0]), vv]);
}
G.CPT_Olsen_layers = (cpt) => {
  const d = columnNum(cpt, 0), C = columnNum(cpt, 5).map((v) => v + 3);
  return [["Layer thickness", "SCN+3"], ...layersOf(d, C)];
};
G.CPT_Robertson_layers = (cpt) => {
  const d = columnNum(cpt, 0), S = columnNum(cpt, 5), Sn = columnNum(cpt, 8);
  const a = layersOf(d, S), b = layersOf(d, Sn), n = Math.max(a.length, b.length), rows = [];
  for (let i = 0; i < n; i++) rows.push([...(a[i] || [NAV, NAV]), ...(b[i] || [NAV, NAV])]);   // HSTACK pads with #N/A
  return [["SBT thickness", "SBT", "SBTIc thickness", "SBTIc"], ...rows];
};
G.CPT_Olsen_SCN_row = (Qt, Fr) => {
  const x = Fr > 0 ? log10(Fr) : -2, LQ = Qt > 0 ? log10(Qt) : -2;
  const m2 = x > 2.7 ? 4.5744 * x ** 3 - 11.185 * x ** 2 + 10.632 * x - 2.8426 : -10;
  const m1 = x > 0.5 ? (x < 0.28 ? -1.7487 * x ** 2 + 1.44 * x + 0.6477 : -0.573 * x ** 2 + 1.9578 * x + 0.4221) : -10;
  const s0 = x > 0.13 ? (x < 0.28 ? 1.0897 * x ** 3 - 0.024 * x ** 2 + 0.5916 * x + 1.3354 : -0.1418 * x ** 2 + 1.2107 * x + 1.198) : -10;
  const s1 = x < 0.32 ? -0.285 * x ** 2 + 0.448 * x + 1.8288 : -0.1418 * x ** 2 + 1.2107 * x + 1.198;
  const s2 = x < 0.42 ? -0.2941 * x ** 2 + 0.2949 * x + 2.3987 : 1.0434 * x + 2.043;
  if (-1 > x || x > 1 || LQ < 0 || LQ > 3) return 7;
  return LQ <= m2 ? 1 : LQ <= m1 ? 2 : LQ <= s0 ? 3 : LQ <= s1 ? 4 : LQ <= s2 ? 5 : 6;
};
const OLSEN_NAMES = ["organic clay and unstable clayey silts", "clays", "silt mixtures", "sand mixtures", "sands", "sands and gravels", "NA"];
const OLSEN_SCN = [-2, -1, 0, 1, 2, 3, 10];
function normalizeCPT(dqs, pa, n, g, gw, dw) {
  return mat(dqs).map((r) => {
    const d = num(r[0]), Q = num(r[1]), F = num(r[2]), sigt = g * d, sige = sigt - gw * (d > dw ? d - dw : 0);
    return [(Q - sigt) / pa * (pa / sige) ** n, F / (Q - sigt) * 100];
  });
}
G.CPT_Olsen_SCN = (dqs, pa, n, g, gw, dw) => normalizeCPT(dqs, pa, n, g, gw, dw).map(([Qn, Fr]) => {
  const i = G.CPT_Olsen_SCN_row(Qn, Fr);
  return [Qn, Fr, OLSEN_SCN[i - 1], OLSEN_NAMES[i - 1]];
});
G.CPT_Robertson_classification_chart = (x_, sel) => {
  const x = log10(x_), c = upper(sel);
  const y = { A: -0.6662 * x ** 3 - 1.7072 * x ** 2 - 1.5868 * x + 0.4627, B: 1.2802 * x ** 3 - 1.2855 * x ** 2 + 1.0234 * x - 0.1389,
    C: 1.1037 * x ** 3 - 0.009 * x ** 2 + 0.6298 * x + 0.7673, D: 1.0659 * x ** 3 - 0.0174 * x ** 2 - 2.6277 * x + 3.3431,
    E: 0.7223 * x ** 4 + 0.6844 * x ** 3 + 0.4039 * x ** 2 + 0.6797 * x + 1.2115, F: 0.2897 * x ** 3 + 0.8671 * x ** 2 + 0.9718 * x + 1.8201,
    G: 1.8669 * x ** 4 + 4.5888 * x ** 3 + 4.4977 * x ** 2 + 2.5417 * x + 2.9667, H: 9.682 * x - 3.7227 }[c];
  return y === undefined ? NAV : 10 ** y;
};
G.CPT_Robertson_SBT_row = (Qt_, Fr_) => {
  const blank = (v) => v === undefined || v === "";
  const Qt = blank(Qt_) ? -1 : num(Qt_), Fr = blank(Fr_) ? -1 : num(Fr_);
  const x = Fr > 0 ? log10(Fr) : -2, LQ = Qt > 0 ? log10(Qt) : -2;
  const a = -0.6662 * x ** 3 - 1.7072 * x ** 2 - 1.5868 * x + 0.4627, b = 1.2802 * x ** 3 - 1.2855 * x ** 2 + 1.0234 * x - 0.1389;
  const c = 1.1037 * x ** 3 - 0.009 * x ** 2 + 0.6298 * x + 0.7673, d = 1.0659 * x ** 3 - 0.0174 * x ** 2 - 2.6277 * x + 3.3431;
  const e = 0.7223 * x ** 4 + 0.6844 * x ** 3 + 0.4039 * x ** 2 + 0.6797 * x + 1.2115, f = 0.2897 * x ** 3 + 0.8671 * x ** 2 + 0.9718 * x + 1.8201;
  const g = 1.8669 * x ** 4 + 4.5888 * x ** 3 + 4.4977 * x ** 2 + 2.5417 * x + 2.9667, h = 9.682 * x - 3.7227;
  if (-1 > x || x > 1 || Qt < 0 || LQ > 3) return 0;
  if (LQ <= a) return 1; if (LQ <= b) return 2; if (LQ > g) return 7;
  if (LQ > d && LQ > h) return 8; if (LQ > d && LQ <= h) return 9;
  if (LQ <= c) return 3; if (LQ <= e) return 4; if (LQ <= f) return 5;
  return 6;
};
const SBT_NAMES = ["Sensitive fine-grained", "Clay - organic soil", "Clays: clay to silty clay", "Silt mixtures: clayey silt & silty clay",
  "Sand mixtures: silty sand to sandy silt", "Sands: clean sands to silty sands", "Dense sand to gravelly sand", "Stiff sand to clayey sand", "Stiff fine-grained", "NA"];
const sbtName = (i) => (i >= 1 && i <= 10 ? SBT_NAMES[i - 1] : NAV);
G.CPT_Robertson_SBT = (dqs, pa, n, g, gw, dw) => normalizeCPT(dqs, pa, n, g, gw, dw).map(([Qn, Fr]) => {
  const S = G.CPT_Robertson_SBT_row(Qn, Fr);
  const Ic = sqrt((3.47 - log10(abs(Qn))) ** 2 + (1.22 + (Fr > 0 ? log10(Fr) : -4)) ** 2), SIc = ROUND(9.2 - 1.877 * Ic, 0);
  return [Qn, Fr, S, sbtName(S), Ic, SIc, sbtName(SIc)];
});
G.CPT_settlement_Iz = (z, B, D, L, qn, g) => {
  const x = 1 - B / L / 2, zm = D + B * x, zl = D + 4 * B * x, ID = 0.2 * x, Im = 0.5 + 0.1 * sqrt(qn / (g * zm));
  return (D <= z && z < zm ? ID + (Im - ID) * (z - D) / (zm - D) : 0) + (zm <= z && z <= zl ? Im - Im * (z - zm) / (zl - zm) : 0);
};
G.CPT_settlement = (zqc, k, B, D, L, qn, g) => {
  const z = columnNum(zqc, 0), qc = columnNum(zqc, 1), za = pairMean(z), qca = pairMean(qc), dz = diffs(z);
  const c = Math.max(1 - g * D / qn / 2, 0.5);
  const rows = za.map((zz, i) => { const Iz = G.CPT_settlement_Iz(zz, B, D, L, qn, g), E = k * qca[i]; return [zz, qca[i], E, Iz, E > 0 ? c * Iz * qn / E : 0]; });
  const s = sum(rows.map((r, i) => r[4] * dz[i]));
  return [["Settlement =", s, "", "", ""], ["depth", "qc", "E", "I", "eps"], ...filter(rows, rows.map((r) => r[3] > 0))];
};
G.CPT_stepProfile = G.stepProfile;

// ---------------------------------------------------------------- lateral piles
G.lateralpile_Broms_cohesionless = (D, phi, g, B, e) => {
  const f = rad(phi), Kp = (1 + sin(f)) / (1 - sin(f)), V = B * D ** 3 * g * Kp / 2 / (D + e), dm = sqrt(6 * V / (B * g * Kp)) / 3;
  return [V, dm, (e + 2 / 3 * dm) * V];
};
G.lateralpile_Broms_cohesive = (V, e, Su, B) => {
  const SuB = 9 * Su * B, f = abs(V) / SuB, Bd = 1.5 * B, D = 2 * sqrt(f * (e + Bd + f / 2)) + Bd + f;
  return [D, (Bd + f + D) / 2, f + Bd, (e + Bd + f / 2) * V];
};
function shearMomentFromReaction(x, px, V) {
  const pxa = pairMean(px), dx = diffs(x);
  const Sx = cumsum([0, ...pxa.map((v, i) => v * dx[i])], V);
  const Sa = pairMean(Sx), mx = cumsum([0, ...Sa.map((v, i) => v * dx[i])]);
  return [["Depth", "Reaction", "Shear", "Moment"], ...x.map((xx, i) => [xx, px[i], Sx[i], -mx[i]])];
}
G.lateralpile_cohesionless_reaction_shear_moment = (D, phi, g, B, e) => {
  const f = rad(phi), Kp = (1 + sin(f)) / (1 - sin(f)), Kps = 3 * g * B * Kp, V = g * B * D ** 3 * Kp / 2 / (D + e);
  const n = 100, ds = -(D + e) / (n - 1);
  const x = sortUnique([0, ...seq(n, e, ds)], true);
  return shearMomentFromReaction(x, x.map((v) => (v > 0 ? 0 : -Kps * v)), V);
};
G.lateralpile_cohesive_reaction_shear_moment = (V, e, Su, B, D_, Dr_) => {
  const SuB = 9 * Su * B, f = abs(V) / SuB, Bd = 1.5 * B;
  const D = D_ === undefined ? 2 * sqrt(f * (e + Bd + f / 2)) + Bd + f : num(D_), Dr = Dr_ === undefined ? (Bd + f + D) / 2 : num(Dr_);
  const n = 100, eps = 0.001, ds = -(D + e) / (n - 1);
  const x = sortUnique([e, 0, -Bd, -Bd - eps, -Dr, -Dr - eps, -D, ...seq(n, e, ds)], true);
  return shearMomentFromReaction(x, x.map((v) => (v > -Bd ? 0 : v < -Dr ? -SuB : SuB)), V);
};

// ---------------------------------------------------------------- axial piles
const alphaFactor = (Sa, a) => {
  const at = -0.0744 * Sa ** 3 + 0.477 * Sa ** 2 - 1.0243 * Sa + 1.0924;
  return a === 1 ? (at < 1 ? at : 1) : a === 2 ? (Sa <= 1.5 ? 0.55 : Sa < 2.5 ? 0.55 - 0.1 * (Sa - 1.5) : 0.45) : 0;
};
G.pile_sidefriction_alpha = (Su, author) => alphaFactor(Su, xmatchNextLarger(lower(author), ["terzaghi", "reese"]));
G.pile_sidefriction_lambda = (D_, ft) => { const d = (ft === undefined ? 1 / 0.3048 : 1) * D_; return 0.091 + 23.37 * (d + 27.05) ** -1.227; };
G.pile_shaft_alpha = (z, Su, pa, p, L, author) => {
  const a = xmatchNextLarger(lower(author), ["terzaghi", "reese"]), aS = Su.map((s) => alphaFactor(s / pa, a) * s);
  const fn = pairMean(aS), dz = diffs(z), zd = drop1(z);
  return [["Ultimate shaft resistance =", p * sum(filter(fn.map((v, i) => v * dz[i]), zd.map((v) => v <= L)))]];
};
G.pile_tip_meyerhof_factor = (phi) => {
  const f = rad(phi), Nq = (1 + sin(f)) / (1 - sin(f)) * exp(5 * PI / 3 * tan(f)), ft = phi === 0 ? 0.1 : f;
  return [Nq, Math.max(9, (Nq - 1) / tan(ft))];
};
G.pile_tip_Meyerhof = (pa, c, phi) => {
  const [Nq, Nc] = G.pile_tip_meyerhof_factor(phi);
  return [["Ultimate tip pressure =", c * Nc + pa * Nq * tan(rad(phi)) / 2]];
};
G.pile_tip_vesic_factor = (phi, Ir) => {
  const f = rad(phi), Nq = 3 / (3 - sin(f)) * exp((PI / 2 - f) * tan(f)) * tan(PI / 4 + f / 2) ** 2 * Ir ** (4 / 3 * sin(f) / (1 + sin(f)));
  return [Nq, phi === 0 ? 4 / 3 * (ln(Ir) + 1) + PI / 2 + 1 : (Nq - 1) / tan(f)];
};
G.pile_tip_vesic = (sv, pa, c, phi_, Dr_) => {
  const phi = phi_ === undefined ? 0 : num(phi_), Dr = Dr_ === undefined ? 0.1 : num(Dr_), f = rad(phi);
  const K0 = 1 - sin(f), sm = (1 + 2 * K0) / 3 * sv, Es = 1000 * Dr * pa, nu = 0.1 + 0.3 * (phi - 25) / 20, Gs = Es / (1 + nu) / 2;
  const Ir = Gs / (c + sm * tan(f)), ev = phi === 0 ? 0 : 0.005 * (1 - (phi - 25) / 20) * sv / pa;
  const Irr = phi === 0 ? 305.06 * c / pa - 12.5 : Ir / (1 + ev * Ir);
  const Nq = phi === 0 ? 1 : 3 / (3 - sin(f)) * exp((PI / 2 - f) * tan(f)) * tan(PI / 4 + f / 2) ** 2 * Irr ** (4 / 3 * sin(f) / (1 + sin(f)));
  const Nc = phi === 0 ? 4 / 3 * (ln(Irr) + 1) + PI / 2 + 1 : (Nq - 1) / tan(f);
  return [["Ultimate pressure =", c * Nc + sm * Nq], ["Nq =", Nq], ["Nc =", Nc], ["Reduced rigidity index Irr =", Irr], ["Mean stress =", sm]];
};
/** Depth grid used by the capacity functions: SORT(UNIQUE(VSTACK(z, extra...))). */
const depthGrid = (z, ...extra) => sortUnique([...z, ...extra]);
G.pile_capacity_alpha = (zSu, pa, a, p, L, author, D_) => {
  const z_ = columnNum(zSu, 0), Su_ = columnNum(zSu, 1), d = D_ === undefined ? 0 : num(D_), Dt = d + L;
  const z = depthGrid(z_, d, Dt), Su = G.interpV(z, z_, Su_);
  if (Su.includes(NAV)) throw new XlErr(NA_CODE, "depth outside the Su profile");
  const au = xmatchNextLarger(lower(author), ["terzaghi", "reese"]);
  const fn = pairMean(Su.map((s) => alphaFactor(s / pa, au) * s)), zd = drop1(z), dz = diffs(z);
  const Qs = p * sum(filter(fn.map((v, i) => v * dz[i]), zd.map((v) => v > d && v <= Dt)));
  const Qt = 9 * G.interp(Dt, z, Su) * a;
  return [["Ultimate pile capacity =", Qt + Qs], ["Ultimate shaft capacity =", Qs], ["Ultimate tip capacity =", Qt]];
};
G.pile_capacity_CPT_Schmertman = (tbl, A, p, L, pile) => {
  const m = mat(tbl), z = columnNum(m, 0), qc = columnNum(m, 1), fs = columnNum(m, 2);
  const k = xmatchNextLarger(lower(pile), ["steel", "concrete", "timber"]);
  if (!k) throw new XlErr(VALUE_CODE, "unknown pile type");
  const soil = m.map((r) => { const s = r.length > 3 ? lower(r[3]) : "sand"; const i = xmatchNextLarger(s, ["clay", "silt", "sand", "gravel"]); return i || 3; });
  const B = 4 * A / p, K = [[0.78, 3.385, 0.1775], [0.83, 1.629, 0.1736], [0.98, 4.23, 0.1775]][k - 1];
  const fsx = z.map((zz, i) => {
    const Ks = K[0] + K[1] / exp(zz / B) ** K[2], fsc = k <= 2 ? (fs[i] < 0.04 ? fs[i] : 0.04) : (fs[i] <= 0.03 ? fs[i] : 0.03 + (fs[i] - 0.03) * 0.294);
    return soil[i] <= 2 ? fsc : Ks * fs[i];
  });
  const fsa = pairMean(fsx), zd = drop1(z), dz = diffs(z);
  const Qs = sum(filter(zd.map((v, i) => v * p * fsa[i] * dz[i] / 8 / B), zd.map((v) => v < 8 * B)))
    + sum(filter(zd.map((v, i) => p * fsa[i] * dz[i]), zd.map((v) => v >= 8 * B && v <= L)));
  const Lt = L - 8 * B, qca = pairMean(qc);
  const qta = sum(filter(qca.map((v, i) => v * dz[i]), zd.map((v) => v >= Lt && v <= L))) / 8 / B;
  const Lb = L + 3.75 * B, sel = zd.map((v) => v >= L && v <= Lb);
  const qx = cumsum(filter(qca.map((v, i) => v * dz[i]), sel)), zx = cumsum(filter(dz, sel));
  const qtb = Math.min(...filter(qx.map((v, i) => v / zx[i]), zx.map((v) => v >= 0.7 * B && v <= 3.75 * B)));
  const Qt = (qta + qtb) * A / 2;
  return [["Ultimate capacity =", Qt + Qs], ["Shaft capacity =", Qs], ["Tip capacity =", Qt], ["Average fs along shaft =", Qs / L / p], ["Average qc around tip =", Qt / A]];
};
G.pile_capacity_CPT_LCPC = (tbl, a, p, L, pile) => {
  const m = mat(tbl), z_ = columnNum(m, 0), qc_ = columnNum(m, 1), soil_ = column(m, 2).map(str);
  const B = 4 * a / p, Dt = 3 * B, Lt = L - 1.5 * B, Lb = Lt + Dt, z = depthGrid(z_, L, Lt, Lb);
  const soil = z.map((v) => xlookup(v, z_, soil_, "", -1)), qc = G.interpV(z, z_, qc_);
  if (qc.includes(NAV)) throw new XlErr(NA_CODE, "depth outside the qc profile");
  const cls = qc.map((q, i) => G.classify_CPT_LCPC(q, soil[i], pile));
  const kc = xlookup(L, z, cls.map((r) => r[1]), "", 1);
  const zd = drop1(z), dz = diffs(z), qca = pairMean(qc);
  const qce = sum(filter(qca.map((v, i) => v * dz[i]), zd.map((v) => v > Lt && v <= Lb))) / Dt;
  const Qt = a * kc * qce;
  const qcs = drop1(qc.map((q, i) => { const qb = q / cls[i][2]; return qb > cls[i][3] ? cls[i][3] : qb; }));
  const qct = sum(filter(qcs.map((v, i) => v * dz[i]), zd.map((v) => v <= L)));
  return [["Pile ultimate capacity =", Qt + qct * p], ["Shaft capacity =", qct * p], ["Tip capacity =", Qt], ["Average qc around tip =", qce], ["Average qs along shaft =", qct / L], ["kc =", kc]];
};
const FELL = {
  1: [[0, 25, 30, 50], [0.15, 0.15, 0.35, 0.35], [3, 3, 30, 30]], 2: [[0, 28, 34, 50], [0.25, 0.25, 0.5, 0.5], [20, 20, 40, 40]],
  3: [[0, 32, 40, 50], [0.3, 0.3, 0.9, 0.9], [30, 30, 150, 150]], 4: [[0, 35, 45, 50], [0.35, 0.35, 0.8, 0.8], [60, 60, 300, 300]],
};
G.pile_capacity_Fellenius = (tbl, a, p, L, D_) => {
  const m = mat(tbl), z_ = columnNum(m, 0), phi_ = columnNum(m, 1), sv_ = columnNum(m, 2), soil_ = column(m, 3).map(str);
  const d = D_ === undefined ? 0 : num(D_), Dt = d + L, z = depthGrid(z_, d, Dt);
  const soil = z.map((v) => xlookup(v, z_, soil_, "", -1)), phi = G.interpV(z, z_, phi_), sv = G.interpV(z, z_, sv_);
  if (phi.includes(NAV) || sv.includes(NAV)) throw new XlErr(NA_CODE, "depth outside the profile");
  const s = soil.map((v) => xmatchNextLarger(lower(v), ["clay", "silt", "sand", "gravel"]));
  const beta = phi.map((f, i) => (FELL[s[i]] ? G.interp(f, FELL[s[i]][0], FELL[s[i]][1]) : 0));
  const Ns = phi.map((f, i) => (FELL[s[i]] ? G.interp(f, FELL[s[i]][0], FELL[s[i]][2]) : 0));
  if (beta.includes(NAV) || Ns.includes(NAV)) throw new XlErr(NA_CODE, "friction angle outside the Fellenius table");
  const Qt = G.interp(Dt, z, Ns) * G.interp(Dt, z, sv) * a;
  const zd = drop1(z), dz = diffs(z), fn = pairMean(beta.map((b, i) => b * sv[i]));
  const Qs = p * sum(filter(fn.map((v, i) => v * dz[i]), zd.map((v) => v > d && v <= Dt)));
  return [["Ultimate pile capacity =", Qt + Qs], ["Shaft capacity =", Qs], ["Tip capacity =", Qt]];
};
const PMT_Q = {
  1: [[0.2, 0.4, 1, 1.4, 5], [0, 0.018, 0.035, 0.037, 0.04]], 2: [[0.2, 0.6, 1, 1.4, 1.7, 5], [0, 0.04, 0.06, 0.07, 0.08, 0.08]],
  3: [[0.2, 0.8, 1.7, 2, 2.4, 5], [0, 0.07, 0.11, 0.118, 0.12, 0.12]], 4: [[0.2, 1, 1.4, 1.8, 2.2, 2.6, 3, 5], [0, 0.09, 0.117, 0.135, 0.15, 0.158, 0.16, 0.16]],
  5: [[0.2, 1.6, 5], [0, 0.16, 0.26]], 6: [[0, 1.4, 5], [0.04, 0.18, 0.3]],
};
G.PMT_shaftFriction = (pl, i) => {
  if (i === 7) return G.interp(pl, [0, 2.6], [0.04, 0.3]);
  return PMT_Q[i] ? G.interp(pl, PMT_Q[i][0], PMT_Q[i][1]) : 0;
};
G.PMT_pile_shaftFriction = (pl, c) => {
  const k = lower(c);
  if (k === "q7") return G.interp(pl, [0, 2.6, 5], [0.04, 0.3, 0.3]);
  const i = { q1: 1, q2: 2, q3: 3, q4: 4, q5: 5, q6: 6 }[k];
  return i ? G.interp(pl, PMT_Q[i][0], PMT_Q[i][1]) : NAV;
};
G.pile_capacity_PMT_LCPC = (tbl, a, p, L, pile) => {
  const m = mat(tbl), z_ = columnNum(m, 0), pl_ = columnNum(m, 1), sv_ = columnNum(m, 2), soil_ = column(m, 3).map(str);
  const z = depthGrid(z_, L), pl = G.interpV(z, z_, pl_), sv = G.interpV(z, z_, sv_);
  if (pl.includes(NAV) || sv.includes(NAV)) throw new XlErr(NA_CODE, "depth outside the profile");
  const soil = z.map((v) => xlookup(v, z_, soil_, "", -1)), cls = pl.map((v, i) => G.classify_PMT_LCPC(v, soil[i], pile));
  const qu = G.interp(L, z, pl.map((v, i) => cls[i][0] * v)) + G.interp(L, z, sv), Qt = qu * a;
  const fs = pl.map((v, i) => G.PMT_shaftFriction(v, cls[i][1]));
  if (fs.includes(NAV)) throw new XlErr(NA_CODE);
  const fsa = pairMean(fs), zd = drop1(z), dz = diffs(z);
  const Qs = p * sum(filter(fsa.map((v, i) => v * dz[i]), zd.map((v) => v >= 0 && v <= L)));
  return [["Ultimate capacity =", Qt + Qs], ["Shaft capacity =", Qs], ["Tip capacity =", Qt]];
};
G.pile_capacity_SPT_Meyerhof = (zN, pa, a, p, L, disp) => {
  const z_ = columnNum(zN, 0), N_ = columnNum(zN, 1), B = 4 * a / p, Dt = 13 * B, Lt = L - 10 * B, Lb = Lt + Dt;
  const z = depthGrid(z_, L, Lt, Lb), N = G.interpV(z, z_, N_);
  if (N.includes(NAV)) throw new XlErr(NA_CODE, "depth outside the N profile");
  const zd = drop1(z), dz = diffs(z), Na = pairMean(N);
  const Ns = sum(filter(Na.map((v, i) => v * dz[i]), zd.map((v) => v <= L)));
  const Qs = (disp === undefined ? 1 : 2) / 100 * pa * Ns * p;
  const Nt = sum(filter(Na.map((v, i) => v * dz[i]), zd.map((v) => v >= Lt && v <= Lb)));
  const Qt = (L / B < 10 ? 0.4 * L / B : 4) * pa * Nt / Dt * a;
  return [["Pile ultimate capacity =", Qt + Qs], ["Shaft capacity =", Qs], ["Tip ultimate capacity =", Qt], ["Average N along shaft =", Ns / L], ["Average N around tip =", Nt / Dt]];
};
G.pile_shaft_lambda = (tbl, p, L, ft) => {
  const z_ = columnNum(tbl, 0), Su_ = columnNum(tbl, 1), sv_ = columnNum(tbl, 2), z = depthGrid(z_, L);
  const Su = G.interpV(z, z_, Su_), sv = G.interpV(z, z_, sv_);
  if (Su.includes(NAV) || sv.includes(NAV)) throw new XlErr(NA_CODE, "depth outside the profile");
  const lm = 0.091 + 23.37 * ((ft === undefined ? 1 / 0.3048 : 1) * L + 27.05) ** -1.227;
  const dz = diffs(z), zd = drop1(z), sel = zd.map((v) => v <= L);
  const Sav = sum(filter(pairMean(Su).map((v, i) => v * dz[i]), sel)), svav = sum(filter(pairMean(sv).map((v, i) => v * dz[i]), sel));
  return [["Pile shaft capacity =", lm * (2 * Sav + svav) * p], ["Lambda =", lm], ["Average undrained shear strength =", Sav / L], ["Average vertical effective stress =", svav / L]];
};

// ---------------------------------------------------------------- pressuremeter (Ménard)
G.PMT_alpha = (soil, Epl) => {
  const s = xmatchNextLarger(lower(soil), ["peat", "clay", "silt", "sand", "gravel", "rock"]);
  const T = { 2: [[0, 7, 9, 16, 1000], [0.5, 0.5, 0.667, 1, 1]], 3: [[0, 5, 8, 14, 1000], [0.5, 0.5, 0.5, 0.667, 0.667]],
    4: [[0, 5, 7, 12, 1000], [0.333, 0.333, 0.5, 0.5, 0.5]], 5: [[0, 6, 10, 1000], [0.25, 0.25, 0.333, 0.333]] }[s];
  if (s === 1) return 1;
  return T ? xlookup(Epl, T[0], T[1], 0, -1) : 0.5;
};
G.PMT_bearing_limitPressure = (z, pl, B, D) => {
  const zmax = D + 3 * B / 2;
  const pld = filter(pl, z.map((v) => v >= D && v <= zmax), [geomean(pl)]), ple = geomean(pld);
  const plu = filter(pl, z.map((v) => v <= D), [ple]);
  let De; try { De = D * (average(plu) / ple); } catch (e) { De = D; }
  return [ple, De];
};
G.PMT_bearing_Menard_k = (H, BL, soil) => {
  const s = upper(soil);
  const k0 = { I: 0.0026 * H ** 3 - 0.0356 * H ** 2 + 0.2241 * H + 0.8, II: 0.006 * H ** 3 - 0.0638 * H ** 2 + 0.3269 * H + 0.8,
    III: 0.0104 * H ** 3 - 0.0922 * H ** 2 + 0.4515 * H + 0.8, IIIA: 0.0108 * H ** 3 - 0.1105 * H ** 2 + 0.5579 * H + 0.8 }[s] || 0;
  const k1 = { I: 0.0295 * H ** 3 - 0.216 * H ** 2 + 0.706 * H + 0.8, II: 0.0272 * H ** 3 - 0.2275 * H ** 2 + 0.8664 * H + 0.8,
    III: 0.0453 * H ** 3 - 0.352 * H ** 2 + 1.3698 * H + 0.8, IIIA: 0.0612 * H ** 3 - 0.457 * H ** 2 + 1.6548 * H + 0.8 }[s] || 0;
  return k0 + (k1 - k0) * BL;
};
G.PMT_bearing_Menard = (z, pl, B, D, L, soil, g) => {
  const [ple, De] = G.PMT_bearing_limitPressure(z, pl, B, D), k = G.PMT_bearing_Menard_k(De / (B / 2), B / L, soil);
  return [["Ultimate pressure =", k * ple + De * g], ["Pressure limit ple=", ple], ["Factor k =", k]];
};
G.PMT_bearing_eurocode_k = (B, L, D, pl, soil_) => {
  const s = xmatchNextLarger(lower(soil_), ["clay silt", "sand gravel", "chalk"]);
  const y = (0.6 + 0.4 * B / L) * D / B;
  if (s === 1) { const xc = G.interp(pl, [0, 0.7, 1.2, 2, 2.5, 100], [0.249, 0.25, 0.35, 0.351, 0.5, 0.501]); if (xc === NAV) return NAV; return 0.8 * (1 + xc * y); }
  if (s === 2) { const xg = G.interp(pl, [0, 0.5, 1, 2, 2.5, 100], [0.349, 0.35, 0.5, 0.501, 0.8, 0.801]); if (xg === NAV) return NAV; return 1 + xg * y; }
  if (s === 3) return 1.3 * (1 + 0.27 * y);
  return 1 + 0.27 * y;
};
G.PMT_bearing_Eurocode = (z, pl, B, D, L, soil, g) => {
  const [ple, De] = G.PMT_bearing_limitPressure(z, pl, B, D), k = G.PMT_bearing_eurocode_k(B, L, De, ple, soil);
  if (k === NAV) throw new XlErr(NA_CODE);
  return [["Ultimate pressure =", k * ple + De * g], ["Pressure limit ple=", ple], ["Factor k =", k]];
};
G.PMT_layers = (B, D) => {
  const xr = [-5, 11, 11, -5].map((v) => v * B), rows = [];
  for (let k = 0; k < 32; k++) rows.push([xr[k % 4], D - B / 2 + B / 2 * (Math.floor(k / 2) + 1)]);
  return rows;
};
G.PMT_settlement_lambda = (LB, opt_) => {
  const o = opt_ === undefined ? "a" : lower(opt_).slice(0, 1);
  const x = [1, 2, 3, 5, 20, 1000000], ok = LB >= 1 && LB <= 1000000;
  const yv = ok ? G.interp(LB, x, [1.1, 1.2, 1.3, 1.4, 1.5, 1.5001]) : 0, yd = ok ? G.interp(LB, x, [1.13, 1.53, 1.78, 2.14, 2.65, 2.6501]) : 0;
  return o === "v" ? yv : o === "d" ? yd : o === "a" ? [yv, yd] : NAV;
};
G.PMT_moduli = (yEpl, B, D) => {
  const y = columnNum(yEpl, 0), E = columnNum(yEpl, 1), pl = columnNum(yEpl, 2);
  const yl = seq(16).map((k) => D - B / 4 + B / 2 * k), Es = G.interpV(yl, y, E), pls = G.interpV(yl, y, pl);
  if (Es.includes(NAV) || pls.includes(NAV)) throw new XlErr(NA_CODE, "profile does not reach D + 7.75 B");
  const plg = geomean(pls), E1 = Es[0], E2 = Es[1], E35 = harmean(Es.slice(2, 5)), E68 = harmean(Es.slice(5, 8)), E916 = harmean(Es.slice(8, 16));
  const Ed = 1 / (0.25 / E1 + 0.3 / E2 + 0.25 / E35 + 0.1 / E68 + 0.1 / E916);
  const lbl = ["Volumetric modulus Ec =", "Deviatoric modulus Ed =", "Limit pressure =", "Modulus E1 =", "E2 =", "E3-5 =", "E6-8 =", "E9-16 ="];
  const v = [E1, Ed, plg, E1, E2, E35, E68, E916];
  return [...lbl.map((l, i) => [l, v[i], " "]), ["Depth", "Modulus", "Limit pressure"], ...yl.map((d, i) => [d, Es[i], pls[i]])];
};
G.PMT_settlement_basic = (q, Ec, Ed, B, LB, ac, ad, B0_) => {
  const B0 = B0_ === undefined ? 1 : num(B0_), lc = G.PMT_settlement_lambda(LB, "v"), ld = G.PMT_settlement_lambda(LB, "d");
  const sc = q * lc * B * ac / 9 / Ec, sd = 2 * q * B0 * (ld * B / B0) ** ad / 9 / Ed;
  return [sc, sd, sc + sd];
};
G.PMT_settlement = (q, yEpl, B, D, L, soil, ft_) => {
  const B0 = ft_ === undefined ? 0.6 : 2, Es = G.PMT_moduli(yEpl, B, D);
  const Ec = Es[0][1], Ed = Es[1][1], plg = Es[2][1], [lc, ld] = G.PMT_settlement_lambda(L / B, "a");
  const ac = G.PMT_alpha(soil, Ec / plg), ad = G.PMT_alpha(soil, Ed / plg);
  const sc = q * lc * B * ac / 9 / Ec, sd = q * 2 * B0 * (ld * B / B0) ** ad / 9 / Ed;
  const lbl = ["Total settlement =", "Volumetric settlement =", "Deviatoric settlement =", "alphac =", "alphad =", "lambdac =", "lambdad =", "Volumetric modulus Ec =", "Deviatoric modulus Ed="];
  return lbl.map((l, i) => [l, [sc + sd, sc, sd, ac, ad, lc, ld, Ec, Ed][i]]);
};

// ---------------------------------------------------------------- footings: settlement, grain size
G.settlement_Schmertman_factor = (z, Imax, B, za, BL_) => {
  const BL = BL_ === undefined ? 0 : num(BL_), I = 0.2 - 0.1 * BL, zb = za + B - B / 2 * BL, zc = za + 4 * B - 2 * B * BL;
  if (za <= z && z < zb) return I + (Imax - I) * (z - za) / (zb - za);
  if (za <= z && z <= zc) return Imax - Imax * (z - zb) / (zc - zb);
  return 0;
};
G.grainsizeanalysis = (d, p) => {
  const Pmin = Math.min(...p), Pmax = Math.max(...p), Dmin = Math.min(...d), Dmax = Math.max(...d);
  const v = (x) => { if (x === NAV) throw new XlErr(NA_CODE); return x; };
  const D10 = Pmin < 10 && 10 < Pmax ? v(G.interp(10, p, d, "logy")) : 0;
  const D30 = Pmin < 30 && 30 < Pmax ? v(G.interp(30, p, d, "logy")) : 0;
  const D60 = Pmin < 60 && 60 < Pmax ? v(G.interp(60, p, d, "logy")) : 0;
  const clay = Dmin <= 0.005 ? v(G.interp(0.005, d, p, "logx")) : 0, silt = Dmin <= 0.075 ? v(G.interp(0.075, d, p, "logx")) : 0;
  const sand = Dmax >= 4.75 ? v(G.interp(4.75, d, p, "logx")) : 100, gravel = Dmax >= 75 ? v(G.interp(75, d, p, "logx")) : 100;
  return [D10, D30, D60, D10 > 0 ? D60 / D10 : 0, D10 > 0 && D60 > 0 ? D30 ** 2 / D10 / D60 : 0, clay, silt - clay, sand - silt, gravel - sand];
};

// ---------------------------------------------------------------- slope stability
G.slope_circle_cohesive = (pline, center, radius, cu, gam, detail) => {
  const AB = G.intersection_polyline_circle(pline, center, radius), A = AB[0], B = AB[AB.length - 1];
  const ta = ATAN2(A[0] - center[0], A[1] - center[1]), tb = ATAN2(B[0] - center[0], B[1] - center[1]);
  const tav = (ta + tb) / 2, theta = abs(ta - tb), arc = radius * theta, resisting = cu * radius * arc;
  const aS = radius ** 2 / 2 * theta, k = 4 * radius / (3 * theta) * sin(theta / 2);
  const cS = [k * cos(tav) + center[0], k * sin(tav) + center[1]];
  let pts = pline.filter(([x]) => (x - A[0]) * (x - B[0]) < 0);
  if (!pts.length) pts = [A];
  const ext = [center, ...sortRows(uniqueRows([...AB, ...pts]), 0)];
  const pg = G.polygon_area_centroid(ext), cE = [pg[0][1], pg[0][2]], aE = pg[1][1];
  const area = aS - aE, c = [(cS[0] * aS - cE[0] * aE) / area, (cS[1] * aS - cE[1] * aE) / area];
  const FS = resisting / (area * gam * (c[0] - center[0]));
  return detail === undefined ? FS : [[FS, ""], [area, arc], c, A, B];
};
G.slope_circle_polyline = (pl, center, radius) => {
  const pts = G.circle(center, radius).slice(1), AB = G.intersection_polyline_polyline(pts, pl, 2);
  const A = [AB[0][1], AB[0][2]], B = [AB[1][1], AB[1][2]];
  const ta = ATAN2(A[0] - center[0], A[1] - center[1]), tb = ATAN2(B[0] - center[0], B[1] - center[1]);
  let top = pl.filter(([x]) => (x - A[0]) * (x - B[0]) < 0);
  if (!top.length) top = [A];
  const spg = pts.filter(([x, y]) => { const t = ATAN2(x - center[0], y - center[1]); return (t - ta) * (t - tb) < 0; });
  const pac = G.polygon_area_centroid([B, ...spg, A, ...top]);
  return [[pac[0][1], pac[0][2], ""], [A[0], A[1], deg(ta)], [B[0], B[1], deg(tb)], [pac[1][1], "", ""]];
};
G.slope_circle_polyline_cohesive = (pline, center, radius, cu, gam, detail) => {
  const X = G.slope_circle_polyline(pline, center, radius), c = [X[0][0], X[0][1]], A = [X[1][0], X[1][1]], B = [X[2][0], X[2][1]];
  const theta = abs(rad(X[1][2]) - rad(X[2][2])), arc = radius * theta, area = X[3][0];
  const FS = cu * radius * arc / (area * gam * abs(c[0] - center[0]));
  return detail === undefined ? FS : [[FS, ""], [area, arc], c, A, B];
};
G.slope_friction_circle_Taylor = (t, o) => (o === 1 ? sqrt(2) * (1 - t ** 2 / PI ** 2) / sqrt(1 + cos(t)) : t / 2 / sin(t / 2));
G.slope_friction_circle_fct = (FS, f, coh, w, kr, r, OD, DM, alpha, Lc) => {
  const fd = atan(tan(f) / FS), Rf = kr * r * sin(fd), beta = 2 * atan((-DM + sqrt(OD ** 2 + DM ** 2 - Rf ** 2)) / (OD + Rf));
  return coh * Lc / (w * sin(beta) / cos(alpha - beta));
};
G.slope_friction_FS = (FS, f, coh, w, kr, r, a, DM, alpha, ab) => {
  const fd = atan(tan(f) / FS), Rf = kr * r * sin(fd), beta = 2 * atan((-DM + sqrt(a ** 2 + DM ** 2 - Rf ** 2)) / (a + Rf));
  return FS - abs(coh * ab / (w * sin(beta) / cos(alpha - beta)));
};
G.slope_friction_circle = (pline, pt, r, coh, phi, g, opt_) => {
  const f = rad(phi), o = opt_ === undefined ? 0 : 1, res = G.slope_circle_polyline(pline, pt, r);
  const ptC = [res[0][0], res[0][1]], a = [res[1][0], res[1][1]], b = [res[2][0], res[2][1]], w = g * res[3][0];
  const ta = ATAN2(a[0] - pt[0], a[1] - pt[1]), tb = ATAN2(b[0] - pt[0], b[1] - pt[1]), theta = abs(ta - tb);
  const alpha = ATAN2(b[0] - a[0], b[1] - a[1]), La = r * theta, Lc = sqrt((b[0] - a[0]) ** 2 + (b[1] - a[1]) ** 2);
  const d = r * La / Lc, OD = ptC[0] - pt[0], DM = d / cos(alpha) - OD * tan(alpha);
  const kr = G.slope_friction_circle_Taylor(theta, o), FSmin = tan(f) / tan(Math.asin(OD / kr / r)) + 0.01;
  const FSf = seq(100, FSmin, 0.02), df = FSf.map((x) => x - G.slope_friction_circle_fct(x, f, coh, w, kr, r, OD, DM, alpha, Lc));
  const FS = G.polyline_subdivide(0, df.map((v, i) => [v, FSf[i]]));
  const fd = ATAN2(FS, tan(f)), Rf = kr * r * abs(sin(fd));
  const beta = 2 * atan((-DM + sqrt(OD ** 2 + DM ** 2 - Rf ** 2)) / (OD + Rf));
  return [[FS, ""], [w, OD], [La, Lc], [w * cos(alpha) / cos(alpha - beta), deg(beta)], [w * sin(beta) / cos(alpha - beta), deg(alpha)],
    [Rf, deg(theta)], ptC, a, b, [OD, DM]];
};
G.slope_logSpiral = (pl, pt, r, phi, coh, g) => {
  const d = tan(rad(phi)), pts = G.logSpiral(pt, r, phi), ab = G.intersection_polyline_polyline(pts, pl);
  if (ab.length < 2) throw new XlErr(NA_CODE, "spiral must cross the ground twice");
  const a = [ab[0][1], ab[0][2]], b = [ab[1][1], ab[1][2]];
  const ra = (a[0] - pt[0]) ** 2 + (a[1] - pt[1]) ** 2, rb = (b[0] - pt[0]) ** 2 + (b[1] - pt[1]) ** 2;
  const ta = ATAN2(a[0] - pt[0], a[1] - pt[1]), tb = ATAN2(b[0] - pt[0], b[1] - pt[1]);
  const IAB = r ** 2 / 2 / d * abs(exp(-2 * ta * d) - exp(-2 * tb * d));
  const top = filter(pl, pl.map(([x]) => (x - a[0]) * (x - b[0]) < 0));
  const spg = filter(pts, pts.map(([x, y]) => { const q = (x - pt[0]) ** 2 + (y - pt[1]) ** 2; return (q - ra) * (q - rb) < 0; }));
  const pac = G.polygon_area_centroid([a, ...spg, b, ...top]), Xc = [pac[0][1], pac[0][2]], dg = Xc[0] - pt[0], w = g * pac[1][1];
  return [[coh * IAB / (w * dg), "", ""], [Xc[0], Xc[1], w], [a[0], a[1], deg(ta)], [b[0], b[1], deg(tb)], [coh * IAB, w * dg, ""]];
};

// ---------------------------------------------------------------- beams and walls
G.shearMoment = (d, p, k) => {
  const x = [], q = [];
  for (let i = 0; i < d.length - 1; i++) for (let j = 0; j < k; j++) { const t = j / k; x.push(d[i] + (d[i + 1] - d[i]) * t); q.push(p[i] + (p[i + 1] - p[i]) * t); }
  x.push(d[d.length - 1]); q.push(p[p.length - 1]);
  const h = diffs(x), V = [0, ...cumsum(h.map((hh, i) => hh * (q[i] + q[i + 1]) / 2))];
  const M = [0, ...cumsum(h.map((hh, i) => V[i] * hh + hh ** 2 * (2 * q[i] + q[i + 1]) / 6))];
  return x.map((xx, i) => [xx, q[i], V[i], M[i]]);
};
const SM_HDR = ["Depth", "Pressure", "Shear", "Moment"];
G.pressure_shear_moment_anchor = (d, p, da, Ta, n_) => {
  const j = Math.min(d.filter((v) => v <= da).length, d.length - 1);
  if (j < 1) throw new XlErr(VALUE_CODE, "anchor above the first depth");
  const pc = p[j - 1] + (da - d[j - 1]) * (p[j] - p[j - 1]) / (d[j] - d[j - 1]);
  const dd = [...d.slice(0, j), da, da, ...d.slice(j)], pp = [...p.slice(0, j), pc, pc, ...p.slice(j)];
  const rows = G.shearMoment(dd, pp, n_ === undefined ? 20 : num(n_));
  const ia = rows.findIndex((r) => r[0] === da);           // XMATCH(da, x): first exact match
  return [SM_HDR, ...rows.map((r, i) => { const after = i > ia; return [r[0], r[1], r[2] + (after ? Ta : 0), r[3] + (after ? Ta * (r[0] - da) : 0)]; })];
};
/* ============================================================================================
 * Part 3: Excel custom functions (generated by tools/make_wrappers.py — edit the spec, not this file).
 * Every argument is declared any[][] so a column passed where a single value is expected
 * spills a column of results, as the LAMBDA versions do.
 * ==========================================================================================*/

/**
 * Library version and release date, as text. Use =GeoEXCEL.version() to check which version is installed.
 * @customfunction VERSION
 * @helpurl https://jpbardet.github.io/geoXL/index.html#VERSION
 * @returns {any[][]} Result (single value, row or table).
 */
function version() {
  return safe(() => G.version());
}

/**
 * Ultimate bearing pressure combining N factors with shape, depth and inclination factors. Inputs: coh, phi (deg), q surcharge, gam unit weight, B, L, D footing width/length/depth, H, V loads, author (terzaghi, prandtl, meyerhof, hansen, vesic, ec7).
 * @customfunction BEARING_CAPACITY
 * @helpurl https://jpbardet.github.io/geoXL/index.html#BEARING_CAPACITY
 * @param {any[][]} coh Cohesion c (kPa)
 * @param {any[][]} phi Friction angle φ (degrees)
 * @param {any[][]} q Surcharge at footing level, γ·D (kPa)
 * @param {any[][]} gam Unit weight of the soil below the footing (kN/m³)
 * @param {any[][]} B Footing width (m)
 * @param {any[][]} L Footing length (m); use a large value for a strip
 * @param {any[][]} D Embedment depth of the footing (m)
 * @param {any[][]} H Horizontal load (kN); 0 for vertical loading
 * @param {any[][]} V Vertical load (kN)
 * @param {any[][]} author Method: "terzaghi", "prandtl", "meyerhof", "hansen", "vesic" or "ec7"
 * @param {any[][]} [A_] Optional effective footing area (m²), default 1
 * @returns {any[][]} Result (single value, row or table).
 */
function bearing_capacity(coh, phi, q, gam, B, L, D, H, V, author, A_) {
  return safeLifted(() => lift((e_coh, e_phi, e_q, e_gam, e_B, e_L, e_D, e_H, e_V, e_author, e_A_) => G.bearing_capacity(num(e_coh), num(e_phi), num(e_q), num(e_gam), num(e_B), num(e_L), num(e_D), num(e_H), num(e_V), e_author, e_A_), [coh, phi, q, gam, B, L, D, H, V, author, A_], []));
}

/**
 * Depth factors {dc; dq; dγ} for the selected bearing method. Inputs: DB = D/B, phi (deg), author.
 * @customfunction BEARING_DEPTHFACTORS
 * @helpurl https://jpbardet.github.io/geoXL/index.html#BEARING_DEPTHFACTORS
 * @param {any[][]} DB Depth ratio D/B
 * @param {any[][]} phi Friction angle φ (degrees)
 * @param {any[][]} author Method: "terzaghi", "prandtl", "meyerhof", "hansen", "vesic" or "ec7"
 * @returns {any[][]} Result (single value, row or table).
 */
function bearing_depthFactors(DB, phi, author) {
  return safe(() => G.bearing_depthFactors(num(DB), num(phi), str(author)));
}

/**
 * Bearing capacity factors {Nc; Nq; Nγ} for the selected method. Inputs: phi (deg), author_.
 * @customfunction BEARING_FACTORS
 * @helpurl https://jpbardet.github.io/geoXL/index.html#BEARING_FACTORS
 * @param {any[][]} phi Friction angle φ (degrees)
 * @param {any[][]} author_ Method: "terzaghi", "prandtl", "meyerhof", "hansen", "vesic" or "ec7"
 * @returns {any[][]} Result (single value, row or table).
 */
function bearing_factors(phi, author_) {
  return safe(() => G.bearing_factors(num(phi), str(author_)));
}

/**
 * Load inclination factors {ic; iq; iγ}. Inputs: H, V horizontal/vertical loads, BL = B/L, phi (deg), coh, author, optional A_ area (default 1).
 * @customfunction BEARING_INCLINATIONFACTORS
 * @helpurl https://jpbardet.github.io/geoXL/index.html#BEARING_INCLINATIONFACTORS
 * @param {any[][]} H Horizontal load (kN)
 * @param {any[][]} V Vertical load (kN)
 * @param {any[][]} BL Shape ratio B/L
 * @param {any[][]} phi Friction angle φ (degrees)
 * @param {any[][]} coh Cohesion c (kPa)
 * @param {any[][]} author Method: "terzaghi", "prandtl", "meyerhof", "hansen", "vesic" or "ec7"
 * @param {any[][]} [A_] Optional footing area (m²), default 1
 * @returns {any[][]} Result (single value, row or table).
 */
function bearing_inclinationFactors(H, V, BL, phi, coh, author, A_) {
  return safe(() => G.bearing_inclinationFactors(num(H), num(V), num(BL), num(phi), num(coh), str(author), opt(A_)));
}

/**
 * Bearing factor Nq for the selected method (row 2 of bearing_factors). Inputs: Phi (deg), author_.
 * @customfunction BEARING_NQ
 * @helpurl https://jpbardet.github.io/geoXL/index.html#BEARING_NQ
 * @param {any[][]} Phi Friction angle φ (degrees)
 * @param {any[][]} author_ Method: "terzaghi", "prandtl", "meyerhof", "hansen", "vesic" or "ec7"
 * @returns {any[][]} Result (single value, row or table).
 */
function bearing_Nq(Phi, author_) {
  return safeLifted(() => lift((e_Phi, e_author_) => G.bearing_Nq(num(e_Phi), e_author_), [Phi, author_], []));
}

/**
 * Shape factors {sc; sq; sγ} for the selected method. Inputs: BL = B/L, phi (deg), author, opt_ (any value selects the alternative Terzaghi sγ).
 * @customfunction BEARING_SHAPEFACTORS
 * @helpurl https://jpbardet.github.io/geoXL/index.html#BEARING_SHAPEFACTORS
 * @param {any[][]} BL Shape ratio B/L (0 for a strip, 1 for a square)
 * @param {any[][]} phi Friction angle φ (degrees)
 * @param {any[][]} author Method: "terzaghi", "prandtl", "meyerhof", "hansen", "vesic" or "ec7"
 * @param {any[][]} [opt_] Optional: any value selects the alternative Terzaghi sγ
 * @returns {any[][]} Result (single value, row or table).
 */
function bearing_shapeFactors(BL, phi, author, opt_) {
  return safe(() => G.bearing_shapeFactors(num(BL), num(phi), str(author), opt(opt_)));
}

/**
 * Skempton Nc for undrained bearing (polynomial fit). Inputs: DB = D/B (capped at 6), BL = B/L.
 * @customfunction BEARING_SKEMPTON
 * @helpurl https://jpbardet.github.io/geoXL/index.html#BEARING_SKEMPTON
 * @param {any[][]} DB Depth ratio D/B (capped at 6)
 * @param {any[][]} BL Shape ratio B/L
 * @returns {any[][]} Result (single value, row or table).
 */
function bearing_Skempton(DB, BL) {
  return safeLifted(() => lift((e_DB, e_BL) => G.bearing_Skempton(num(e_DB), num(e_BL)), [DB, BL], []));
}

/**
 * AASHTO soil classification with group index, as text. Inputs: % passing No.10, No.40, No.200 (enter 40 for 40%), LL_, PL_.
 * @customfunction CLASSIFY_AASHTO
 * @helpurl https://jpbardet.github.io/geoXL/index.html#CLASSIFY_AASHTO
 * @param {any[][]} pf10_ Percent passing No. 10 sieve (enter 40 for 40 %)
 * @param {any[][]} pf40_ Percent passing No. 40 sieve (%)
 * @param {any[][]} pf200_ Percent passing No. 200 sieve (%)
 * @param {any[][]} [LL_] Liquid limit LL (%), optional
 * @param {any[][]} [PL_] Plastic limit PL (%), optional
 * @returns {any[][]} Result (single value, row or table).
 */
function classify_AASHTO(pf10_, pf40_, pf200_, LL_, PL_) {
  return safeLifted(() => lift((e_pf10_, e_pf40_, e_pf200_, e_LL_, e_PL_) => G.classify_AASHTO(num(e_pf10_), num(e_pf40_), num(e_pf200_), e_LL_, e_PL_), [pf10_, pf40_, pf200_, LL_, PL_], []));
}

/**
 * ASTM D2487 group symbols and group names, 100 rows {symbol, name} (was the ASTMtable name on the ASTMnames sheet). Use as the list argument of classify_ASTM.
 * @customfunction ASTM_TABLE
 * @helpurl https://jpbardet.github.io/geoXL/index.html#ASTM_TABLE
 * @returns {any[][]} Result (single value, row or table).
 */
function astm_table() {
  return safe(() => G.astm_table());
}

/**
 * USCS symbol with ASTM group name from a two-column lookup table (symbol | descriptions). Inputs: list (or "" for the built-in astm_table), % passing No.4 and No.200, optional Cu, Cc, LL, PL, organic.
 * @customfunction CLASSIFY_ASTM
 * @helpurl https://jpbardet.github.io/geoXL/index.html#CLASSIFY_ASTM
 * @param {any[][]} list Two-column table {symbol, group name}; enter "" to use the built-in ASTM_TABLE
 * @param {any[][]} pf4_ Percent passing No. 4 sieve (%)
 * @param {any[][]} pf200_ Percent passing No. 200 sieve (%)
 * @param {any[][]} [Cu_] Coefficient of uniformity Cu = D60/D10, optional
 * @param {any[][]} [Cc_] Coefficient of curvature Cc = D30²/(D10·D60), optional
 * @param {any[][]} [LL_] Liquid limit LL (%), optional
 * @param {any[][]} [PL_] Plastic limit PL (%), optional
 * @param {any[][]} [organic_] Organic indicator: omit or 0 = inorganic, 1 = organic
 * @returns {any[][]} Result (single value, row or table).
 */
function classify_ASTM(list, pf4_, pf200_, Cu_, Cc_, LL_, PL_, organic_) {
  const list_c = mat(list);
  return safeLifted(() => lift((e_list, e_pf4_, e_pf200_, e_Cu_, e_Cc_, e_LL_, e_PL_, e_organic_) => G.classify_ASTM(e_list, num(e_pf4_), num(e_pf200_), e_Cu_, e_Cc_, e_LL_, e_PL_, e_organic_), [list_c, pf4_, pf200_, Cu_, Cc_, LL_, PL_, organic_], [0]));
}

/**
 * Fine-grained soil symbol from the plasticity chart (CL, CH, ML, MH, OL, OH, CL & ML). Inputs: LL, PL in %, optional organic (0 = inorganic).
 * @customfunction CLASSIFY_PLASTICITY
 * @helpurl https://jpbardet.github.io/geoXL/index.html#CLASSIFY_PLASTICITY
 * @param {any[][]} LL Liquid limit LL (%)
 * @param {any[][]} PL Plastic limit PL (%)
 * @param {any[][]} [organic] Organic indicator: omit or 0 = inorganic, 1 = organic
 * @returns {any[][]} Result (single value, row or table).
 */
function classify_Plasticity(LL, PL, organic) {
  return safeLifted(() => lift((e_LL, e_PL, e_organic) => G.classify_Plasticity(num(e_LL), num(e_PL), e_organic), [LL, PL, organic], []));
}

/**
 * USCS group symbol from gradation and plasticity. Inputs: % passing No.4 and No.200, optional Cu = D60/D10, Cc = D30²/(D10·D60), LL, PL, organic.
 * @customfunction CLASSIFY_USC
 * @helpurl https://jpbardet.github.io/geoXL/index.html#CLASSIFY_USC
 * @param {any[][]} pf4_ Percent passing No. 4 sieve (%)
 * @param {any[][]} pf200_ Percent passing No. 200 sieve (%)
 * @param {any[][]} [Cu_] Coefficient of uniformity Cu = D60/D10, optional
 * @param {any[][]} [Cc_] Coefficient of curvature Cc = D30²/(D10·D60), optional
 * @param {any[][]} [LL_] Liquid limit LL (%), optional
 * @param {any[][]} [PL_] Plastic limit PL (%), optional
 * @param {any[][]} [organic_] Organic indicator: omit or 0 = inorganic, 1 = organic
 * @returns {any[][]} Result (single value, row or table).
 */
function classify_USC(pf4_, pf200_, Cu_, Cc_, LL_, PL_, organic_) {
  return safeLifted(() => lift((e_pf4_, e_pf200_, e_Cu_, e_Cc_, e_LL_, e_PL_, e_organic_) => G.classify_USC(num(e_pf4_), num(e_pf200_), e_Cu_, e_Cc_, e_LL_, e_PL_, e_organic_), [pf4_, pf200_, Cu_, Cc_, LL_, PL_, organic_], []));
}

/**
 * Square-root-of-time analysis of a consolidation test: Cv, d0, d100, drainage length and compression ratios. Inputs: time and dial-reading columns, fitting window ta–tb, initial height hi.
 * @customfunction CONSOLIDATION_TEST
 * @helpurl https://jpbardet.github.io/geoXL/index.html#CONSOLIDATION_TEST
 * @param {any[][]} Tarray Column of elapsed times (min)
 * @param {any[][]} Xarray Column of dial readings (same rows as Tarray)
 * @param {any[][]} ta Start of the straight-line fitting window (same unit as Tarray)
 * @param {any[][]} tb End of the fitting window
 * @param {any[][]} hi Initial specimen height (same unit as the readings)
 * @returns {any[][]} Result (single value, row or table).
 */
function consolidation_test(Tarray, Xarray, ta, tb, hi) {
  return safe(() => G.consolidation_test(colNum(Tarray), colNum(Xarray), num(ta), num(tb), num(hi)));
}

/**
 * Time factor Tv from average degree of consolidation U (fraction, 0 ≤ U < 1), approximate formula.
 * @customfunction CONSOLIDATION_TV
 * @helpurl https://jpbardet.github.io/geoXL/index.html#CONSOLIDATION_TV
 * @param {any[][]} U Average degree of consolidation (fraction, 0 ≤ U < 1)
 * @returns {any[][]} Result (single value, row or table).
 */
function consolidation_Tv(U) {
  return safeLifted(() => lift((e_U) => G.consolidation_Tv(num(e_U)), [U], []));
}

/**
 * Average degree of consolidation U (fraction) from time factor tv by Fourier series. case_: uniform (default), triangular_top, triangular_bottom, vshape, saw, sine, parabolic, circular.
 * @customfunction CONSOLIDATION_U
 * @helpurl https://jpbardet.github.io/geoXL/index.html#CONSOLIDATION_U
 * @param {any[][]} tv Time factor Tv (single value or column)
 * @param {any[][]} [case_] Initial excess pore pressure shape: "uniform" (default), "triangular_top", "triangular_bottom", "vshape", "saw", "sine", "parabolic", "circular"
 * @returns {any[][]} Result (single value, row or table).
 */
function consolidation_U(tv, case_) {
  return safe(() => G.consolidation_U(colNum(tv), opt(case_)));
}

/**
 * Excess pore pressure ratio u/u0 along normalized depth at time factor tv. Inputs: Xarray (z/H column), tv, case_ (same cases as consolidation_U).
 * @customfunction CONSOLIDATION_UX
 * @helpurl https://jpbardet.github.io/geoXL/index.html#CONSOLIDATION_UX
 * @param {any[][]} Xarray Column of normalized depths z/H
 * @param {any[][]} tv Time factor Tv
 * @param {any[][]} [case_] Initial pore pressure shape (same choices as consolidation_U)
 * @returns {any[][]} Result (single value, row or table).
 */
function consolidation_Ux(Xarray, tv, case_) {
  return safe(() => G.consolidation_Ux(colNum(Xarray), num(tv), opt(case_)));
}

/**
 * Average degree of consolidation U (fraction) from tv, closed-form approximation.
 * @customfunction CONSOLIDATION_U_APPROXIMATE
 * @helpurl https://jpbardet.github.io/geoXL/index.html#CONSOLIDATION_U_APPROXIMATE
 * @param {any[][]} tv Time factor Tv
 * @returns {any[][]} Result (single value, row or table).
 */
function consolidation_U_approximate(tv) {
  return safeLifted(() => lift((e_tv) => G.consolidation_U_approximate(num(e_tv)), [tv], []));
}

/**
 * Average consolidation U for a load applied linearly over construction time factor tc. Inputs: tv, tc.
 * @customfunction CONSOLIDATION_U_VARIABLE
 * @helpurl https://jpbardet.github.io/geoXL/index.html#CONSOLIDATION_U_VARIABLE
 * @param {any[][]} tv Time factor Tv
 * @param {any[][]} tc Time factor at the end of construction (load ramps up linearly until tc)
 * @returns {any[][]} Result (single value, row or table).
 */
function consolidation_U_variable(tv, tc) {
  return safeLifted(() => lift((e_tv, e_tc) => G.consolidation_U_variable(num(e_tv), num(e_tc)), [tv, tc], []));
}

/**
 * Cosine of an angle in degrees.
 * @customfunction COSD
 * @helpurl https://jpbardet.github.io/geoXL/index.html#COSD
 * @param {any[][]} X Angle (degrees)
 * @returns {any[][]} Result (single value, row or table).
 */
function cosd(X) {
  return safeLifted(() => lift((e_X) => G.cosd(num(e_X)), [X], []));
}

/**
 * Coulomb trial wedge with cohesion: {W; R; P; Cs; Cw; z0; K}. Inputs: t wedge angle (radians), phi, alpha wall, beta slope, delta (deg), g unit weight, h wall height, optional coh_.
 * @customfunction COULOMB_COHESIVEFORCES
 * @helpurl https://jpbardet.github.io/geoXL/index.html#COULOMB_COHESIVEFORCES
 * @param {any[][]} t Trial wedge (slip) angle (radians)
 * @param {any[][]} phi Friction angle φ (degrees)
 * @param {any[][]} alpha Wall inclination α (degrees; 90 = vertical)
 * @param {any[][]} beta Backfill slope β (degrees)
 * @param {any[][]} delta Wall friction angle δ (degrees)
 * @param {any[][]} g Unit weight γ (kN/m³)
 * @param {any[][]} h Wall height H (m)
 * @param {any[][]} [coh_] Cohesion c (kPa), optional, default 0
 * @returns {any[][]} Result (single value, row or table).
 */
function Coulomb_cohesiveforces(t, phi, alpha, beta, delta, g, h, coh_) {
  return safe(() => G.Coulomb_cohesiveforces(num(t), num(phi), num(alpha), num(beta), num(delta), num(g), num(h), opt(coh_)));
}

/**
 * Coulomb trial wedge forces {W; R; P; K} for slip angle t (radians). Inputs: phi, alpha wall (90 = vertical), beta slope, delta (deg), g, h.
 * @customfunction COULOMB_FORCES
 * @helpurl https://jpbardet.github.io/geoXL/index.html#COULOMB_FORCES
 * @param {any[][]} t Trial wedge (slip) angle (radians)
 * @param {any[][]} phi Friction angle φ (degrees)
 * @param {any[][]} alpha Wall inclination α (degrees; 90 = vertical)
 * @param {any[][]} beta Backfill slope β (degrees)
 * @param {any[][]} delta Wall friction angle δ (degrees)
 * @param {any[][]} g Unit weight γ (kN/m³)
 * @param {any[][]} h Wall height H (m)
 * @returns {any[][]} Result (single value, row or table).
 */
function Coulomb_forces(t, phi, alpha, beta, delta, g, h) {
  return safe(() => G.Coulomb_forces(num(t), num(phi), num(alpha), num(beta), num(delta), num(g), num(h)));
}

/**
 * Coulomb earth pressure coefficient. Inputs: phi, alpha, beta, delta (deg); passive: omit for Ka, any value for Kp.
 * @customfunction COULOMB_K
 * @helpurl https://jpbardet.github.io/geoXL/index.html#COULOMB_K
 * @param {any[][]} phi Friction angle φ (degrees)
 * @param {any[][]} alpha Wall inclination α (degrees; 90 = vertical)
 * @param {any[][]} beta Backfill slope β (degrees)
 * @param {any[][]} delta Wall friction angle δ (degrees)
 * @param {any[][]} [passive] Omit for active Ka; any value gives passive Kp
 * @returns {any[][]} Result (single value, row or table).
 */
function coulomb_K(phi, alpha, beta, delta, passive) {
  return safeLifted(() => lift((e_phi, e_alpha, e_beta, e_delta, e_passive) => G.coulomb_K(num(e_phi), num(e_alpha), num(e_beta), num(e_delta), e_passive), [phi, alpha, beta, delta, passive], []));
}

/**
 * Critical Coulomb wedge by golden-section search: {slip angle (deg); W; R; P; K}. Inputs: a_, b_ (any values on first call), phi, alpha, beta, delta, g, h.
 * @customfunction COULOMB_MAXIMUM
 * @helpurl https://jpbardet.github.io/geoXL/index.html#COULOMB_MAXIMUM
 * @param {any[][]} a_ Lower search bound (any value on the first call)
 * @param {any[][]} b_ Upper search bound (any value on the first call)
 * @param {any[][]} phi Friction angle φ (degrees)
 * @param {any[][]} alpha Wall inclination α (degrees; 90 = vertical)
 * @param {any[][]} beta Backfill slope β (degrees)
 * @param {any[][]} delta Wall friction angle δ (degrees)
 * @param {any[][]} g Unit weight γ (kN/m³)
 * @param {any[][]} h Wall height H (m)
 * @param {any[][]} [iter_] Leave empty (internal iteration counter)
 * @returns {any[][]} Result (single value, row or table).
 */
function Coulomb_maximum(a_, b_, phi, alpha, beta, delta, g, h, iter_) {
  return safe(() => G.Coulomb_maximum(num(a_), num(b_), num(phi), num(alpha), num(beta), num(delta), num(g), num(h), opt(iter_)));
}

/**
 * Critical Coulomb slip angle (deg), closed form. Inputs: phi, alpha, beta, delta (deg).
 * @customfunction COULOMB_THETA
 * @helpurl https://jpbardet.github.io/geoXL/index.html#COULOMB_THETA
 * @param {any[][]} phi Friction angle φ (degrees)
 * @param {any[][]} alpha Wall inclination α (degrees; 90 = vertical)
 * @param {any[][]} beta Backfill slope β (degrees)
 * @param {any[][]} delta Wall friction angle δ (degrees)
 * @returns {any[][]} Result (single value, row or table).
 */
function coulomb_theta(phi, alpha, beta, delta) {
  return safeLifted(() => lift((e_phi, e_alpha, e_beta, e_delta) => G.coulomb_theta(num(e_phi), num(e_alpha), num(e_beta), num(e_delta)), [phi, alpha, beta, delta], []));
}

/**
 * Intersections {x, y} of a sampled curve with the line y = slope·x + intercept. Inputs: Xarray, Yarray columns, slope, intercept.
 * @customfunction CUTLINE
 * @helpurl https://jpbardet.github.io/geoXL/index.html#CUTLINE
 * @param {any[][]} Xarray Column of x values of the curve
 * @param {any[][]} Yarray Column of y values of the curve
 * @param {any[][]} slope Slope of the line
 * @param {any[][]} intercept Intercept of the line
 * @returns {any[][]} Result (single value, row or table).
 */
function cutline(Xarray, Yarray, slope, intercept) {
  return safe(() => G.cutline(colNum(Xarray), colNum(Yarray), num(slope), num(intercept)));
}

/**
 * Normalized radius r/z of a circular load giving vertical stress ratio s (0 ≤ s < 1). Used to build influence charts.
 * @customfunction ELASTIC_CIRCULAR_RADIUS
 * @helpurl https://jpbardet.github.io/geoXL/index.html#ELASTIC_CIRCULAR_RADIUS
 * @param {any[][]} s Stress ratio Δσz/q under the centre (0 ≤ s < 1)
 * @returns {any[][]} Result (single value, row or table).
 */
function elastic_circular_radius(s) {
  return safeLifted(() => lift((e_s) => G.elastic_circular_radius(num(e_s)), [s], []));
}

/**
 * Stress ratios {sz, sr} under the centre of a uniformly loaded circle. Inputs: rValue load radius, z depth, optional nu_ (default 0.3). Multiply by the load pressure.
 * @customfunction ELASTIC_CIRCULAR_STRESS
 * @helpurl https://jpbardet.github.io/geoXL/index.html#ELASTIC_CIRCULAR_STRESS
 * @param {any[][]} rValue Radius of the loaded circle (m)
 * @param {any[][]} z Depth below the centre (m)
 * @param {any[][]} [nu_] Poisson's ratio, optional, default 0.3
 * @returns {any[][]} Result (single value, row or table).
 */
function elastic_circular_stress(rValue, z, nu_) {
  return safeLifted(() => liftRows((e_rValue, e_z, e_nu_) => G.elastic_circular_stress(num(e_rValue), num(e_z), e_nu_), [rValue, z, nu_], []));
}

/**
 * Vertical stress influence factor under an embankment (Osterberg). Inputs: b1z = b1/z, b2z = b2/z (b2z ≠ 0).
 * @customfunction ELASTIC_EMBANKMENT
 * @helpurl https://jpbardet.github.io/geoXL/index.html#ELASTIC_EMBANKMENT
 * @param {any[][]} b1z Ratio b1/z (half crest width over depth)
 * @param {any[][]} b2z Ratio b2/z (horizontal length of the side slope over depth, ≠ 0)
 * @returns {any[][]} Result (single value, row or table).
 */
function elastic_embankment(b1z, b2z) {
  return safeLifted(() => lift((e_b1z, e_b2z) => G.elastic_embankment(num(e_b1z), num(e_b2z)), [b1z, b2z], []));
}

/**
 * Janbu μ0 settlement factor from digitized chart. Inputs: DB = D/B, LB = L/B (1, 2, 5, 10, 20, 50, 100 or 200).
 * @customfunction ELASTIC_JANBU_MU0
 * @helpurl https://jpbardet.github.io/geoXL/index.html#ELASTIC_JANBU_MU0
 * @param {any[][]} DB Depth ratio D/B
 * @param {any[][]} LB Length ratio L/B: 1, 2, 5, 10, 20, 50, 100 or 200
 * @returns {any[][]} Result (single value, row or table).
 */
function elastic_Janbu_mu0(DB, LB) {
  return safeLifted(() => lift((e_DB, e_LB) => G.elastic_Janbu_mu0(num(e_DB), num(e_LB)), [DB, LB], []));
}

/**
 * Janbu μ1 settlement factor from digitized chart. Inputs: LB (chart abscissa), DB curve selector ("circle", 1, 2, 5, 10, 20, 50 or 100).
 * @customfunction ELASTIC_JANBU_MU1
 * @helpurl https://jpbardet.github.io/geoXL/index.html#ELASTIC_JANBU_MU1
 * @param {any[][]} LB Chart abscissa (H/B)
 * @param {any[][]} DB Curve: "circle" or L/B = 1, 2, 5, 10, 20, 50, 100
 * @returns {any[][]} Result (single value, row or table).
 */
function elastic_Janbu_mu1(LB, DB) {
  return safeLifted(() => lift((e_LB, e_DB) => G.elastic_Janbu_mu1(num(e_LB), e_DB), [LB, DB], []));
}

/**
 * Stresses {sx, sy, sxy} per unit vertical line load (Flamant). Inputs: x horizontal offset, y depth. Multiply by the line load.
 * @customfunction ELASTIC_LINELOAD
 * @helpurl https://jpbardet.github.io/geoXL/index.html#ELASTIC_LINELOAD
 * @param {any[][]} x Horizontal offset from the load (m)
 * @param {any[][]} y Depth (m)
 * @returns {any[][]} Result (single value, row or table).
 */
function elastic_lineload(x, y) {
  return safeLifted(() => liftRows((e_x, e_y) => G.elastic_lineload(num(e_x), num(e_y)), [x, y], []));
}

/**
 * Stresses {sx, sy, sxy} per unit horizontal line load. Inputs: x horizontal offset, y depth. Multiply by the line load.
 * @customfunction ELASTIC_LINELOAD_HORIZONTAL
 * @helpurl https://jpbardet.github.io/geoXL/index.html#ELASTIC_LINELOAD_HORIZONTAL
 * @param {any[][]} x Horizontal offset from the load (m)
 * @param {any[][]} y Depth (m)
 * @returns {any[][]} Result (single value, row or table).
 */
function elastic_lineload_horizontal(x, y) {
  return safeLifted(() => liftRows((e_x, e_y) => G.elastic_lineload_horizontal(num(e_x), num(e_y)), [x, y], []));
}

/**
 * Newmark influence factor under the corner of a uniformly loaded rectangle. Inputs: m = B/z, n = L/z.
 * @customfunction ELASTIC_NEWMARK_FACTOR
 * @helpurl https://jpbardet.github.io/geoXL/index.html#ELASTIC_NEWMARK_FACTOR
 * @param {any[][]} m Ratio B/z (rectangle side over depth)
 * @param {any[][]} n Ratio L/z (other side over depth)
 * @returns {any[][]} Result (single value, row or table).
 */
function elastic_Newmark_factor(m, n) {
  return safeLifted(() => lift((e_m, e_n) => G.elastic_Newmark_factor(num(e_m), num(e_n)), [m, n], []));
}

/**
 * Vertical stress factor at any point (x_, y_, z_) under a loaded rectangle by superposition of corner factors. pt: 4 corners, opposite corners in rows 1 and 3.
 * @customfunction ELASTIC_NEWMARK_STRESS
 * @helpurl https://jpbardet.github.io/geoXL/index.html#ELASTIC_NEWMARK_STRESS
 * @param {any[][]} x_ x coordinate of the point
 * @param {any[][]} y_ y coordinate of the point
 * @param {any[][]} z_ Depth of the point
 * @param {any[][]} pt Rectangle corners (4 rows × 2 columns, opposite corners in rows 1 and 3)
 * @returns {any[][]} Result (single value, row or table).
 */
function elastic_Newmark_stress(x_, y_, z_, pt) {
  const pt_c = pts2(pt);
  return safeLifted(() => lift((e_x_, e_y_, e_z_, e_pt) => G.elastic_Newmark_stress(num(e_x_), num(e_y_), num(e_z_), e_pt), [x_, y_, z_, pt_c], [3]));
}

/**
 * Stresses {sz, sr, st, srz} per unit point load (Boussinesq). Inputs: rValue radial offset, z depth, optional nu_ (default 0.3).
 * @customfunction ELASTIC_POINTLOAD
 * @helpurl https://jpbardet.github.io/geoXL/index.html#ELASTIC_POINTLOAD
 * @param {any[][]} rValue Horizontal distance from the load (m)
 * @param {any[][]} z Depth (m)
 * @param {any[][]} [nu_] Poisson's ratio, optional, default 0.3
 * @returns {any[][]} Result (single value, row or table).
 */
function elastic_pointload(rValue, z, nu_) {
  return safeLifted(() => liftRows((e_rValue, e_z, e_nu_) => G.elastic_pointload(num(e_rValue), num(e_z), e_nu_), [rValue, z, nu_], []));
}

/**
 * Vertical stress factor under a point (x, y) at depth z for a uniformly loaded polygon, by integration over 1° sectors. pts: polygon vertices (two columns).
 * @customfunction ELASTIC_SECTOR
 * @helpurl https://jpbardet.github.io/geoXL/index.html#ELASTIC_SECTOR
 * @param {any[][]} x x coordinate of the point
 * @param {any[][]} y y coordinate of the point
 * @param {any[][]} z Depth
 * @param {any[][]} pts Polygon vertices (two columns x, y)
 * @returns {any[][]} Result (single value, row or table).
 */
function elastic_sector(x, y, z, pts) {
  const pts_c = pts2(pts);
  return safeLifted(() => lift((e_x, e_y, e_z, e_pts) => G.elastic_sector(num(e_x), num(e_y), num(e_z), e_pts), [x, y, z, pts_c], [3]));
}

/**
 * Helper for elastic_sector: contribution of one ray. Input: row with angle (radians) followed by flattened polygon coordinates.
 * @customfunction ELASTIC_SECTOR_ROW
 * @helpurl https://jpbardet.github.io/geoXL/index.html#ELASTIC_SECTOR_ROW
 * @param {any[][]} txy Row: ray angle (radians) followed by the flattened polygon coordinates
 * @returns {any[][]} Result (single value, row or table).
 */
function elastic_sector_row(txy) {
  return safe(() => (() => { const v = colNum(txy), poly = []; for (let i = 1; i + 1 < v.length; i += 2) poly.push([v[i], v[i + 1]]); return G.elastic_sector_rowValue(v[0], poly); })());
}

/**
 * Vertical stress contribution of an annular sector between radii ra and rb at depth z. dt: angular weight (sector angle / 2π).
 * @customfunction ELASTIC_SECTOR_SLICE
 * @helpurl https://jpbardet.github.io/geoXL/index.html#ELASTIC_SECTOR_SLICE
 * @param {any[][]} ra Inner radius
 * @param {any[][]} rb Outer radius
 * @param {any[][]} z Depth
 * @param {any[][]} dt Angular weight (sector angle / 2π)
 * @returns {any[][]} Result (single value, row or table).
 */
function elastic_sector_slice(ra, rb, z, dt) {
  return safeLifted(() => lift((e_ra, e_rb, e_z, e_dt) => G.elastic_sector_slice(num(e_ra), num(e_rb), num(e_z), num(e_dt)), [ra, rb, z, dt], []));
}

/**
 * Stress ratios {sy, sx, sxy} under a uniform strip load. Inputs: B full width, x offset from centre, y depth. Multiply by the pressure.
 * @customfunction ELASTIC_STRIPLOAD
 * @helpurl https://jpbardet.github.io/geoXL/index.html#ELASTIC_STRIPLOAD
 * @param {any[][]} B Full width of the strip (m)
 * @param {any[][]} x Horizontal offset from the strip centre (m)
 * @param {any[][]} y Depth (m)
 * @returns {any[][]} Result (single value, row or table).
 */
function elastic_stripload(B, x, y) {
  return safeLifted(() => liftRows((e_B, e_x, e_y) => G.elastic_stripload(num(e_B), num(e_x), num(e_y)), [B, x, y], []));
}

/**
 * Stress ratios {sy, sx, sxy} under a triangular strip load. Inputs: B loaded width, x offset, y depth. Multiply by the maximum pressure.
 * @customfunction ELASTIC_TRIANGLE
 * @helpurl https://jpbardet.github.io/geoXL/index.html#ELASTIC_TRIANGLE
 * @param {any[][]} B Loaded width (m)
 * @param {any[][]} x Horizontal offset (m)
 * @param {any[][]} y Depth (m)
 * @returns {any[][]} Result (single value, row or table).
 */
function elastic_triangle(B, x, y) {
  return safeLifted(() => liftRows((e_B, e_x, e_y) => G.elastic_triangle(num(e_B), num(e_x), num(e_y)), [B, x, y], []));
}

/**
 * Interpolate y at X from (Xa, Ya); X may be a column. Xa may be increasing or decreasing. mode: "linear" (default), "logx", "logy", "logxy". outside: omitted = #N/A, "extrapolate", "clamp".
 * @customfunction INTERPOLATE
 * @helpurl https://jpbardet.github.io/geoXL/index.html#INTERPOLATE
 * @param {any[][]} X Value(s) at which to interpolate
 * @param {any[][]} Xa Column of known x values (increasing or decreasing)
 * @param {any[][]} Ya Column of known y values
 * @param {any[][]} [mode] "linear" (default), "logx", "logy" or "logxy"
 * @param {any[][]} [outside] Omit for #N/A outside the table, "extrapolate" or "clamp"
 * @returns {any[][]} Result (single value, row or table).
 */
function Interpolate(X, Xa, Ya, mode, outside) {
  const Xa_c = colNum(Xa);
  const Ya_c = colNum(Ya);
  const mode_c = opt(mode);
  const outside_c = opt(outside);
  return safeLifted(() => lift((e_X, e_Xa, e_Ya, e_mode, e_outside) => G.interp(num(e_X), e_Xa, e_Ya, e_mode, e_outside), [X, Xa_c, Ya_c, mode_c, outside_c], [1, 2, 3, 4]));
}

/**
 * Intersections {x, y} of a polygon with the infinite line through pt at angle t (radians). Returns {0, 0} if none.
 * @customfunction INTERSECTION_POLYGON_LINE
 * @helpurl https://jpbardet.github.io/geoXL/index.html#INTERSECTION_POLYGON_LINE
 * @param {any[][]} polygon Polygon vertices (two columns x, y)
 * @param {any[][]} pt Point {x, y} on the line
 * @param {any[][]} t Line angle (radians)
 * @returns {any[][]} Result (single value, row or table).
 */
function intersection_polygon_line(polygon, pt, t) {
  return safe(() => G.intersection_polygon_line(pts2(polygon), point(pt), num(t)));
}

/**
 * Sorted distances from pt to the polygon boundary along a ray at angle t (radians).
 * @customfunction INTERSECTION_POLYGON_RAY
 * @helpurl https://jpbardet.github.io/geoXL/index.html#INTERSECTION_POLYGON_RAY
 * @param {any[][]} polygon Polygon vertices (two columns x, y)
 * @param {any[][]} pt Ray origin {x, y}
 * @param {any[][]} t Ray angle (radians)
 * @returns {any[][]} Result (single value, row or table).
 */
function intersection_polygon_ray(polygon, pt, t) {
  return safe(() => G.intersection_polygon_ray(pts2(polygon), point(pt), num(t)));
}

/**
 * Teaching example of golden-section minimization of x³ + 5x² − 10x on [A, B].
 * @customfunction MINIMUM
 * @helpurl https://jpbardet.github.io/geoXL/index.html#MINIMUM
 * @param {any[][]} A Lower bound of the search interval
 * @param {any[][]} B Upper bound
 * @param {any[][]} [iter_] Leave empty (internal counter)
 * @returns {any[][]} Result (single value, row or table).
 */
function minimum(A, B, iter_) {
  return safeLifted(() => lift((e_A, e_B, e_iter_) => G.minimum(num(e_A), num(e_B), e_iter_), [A, B, iter_], []));
}

/**
 * Alias of polygon_area_centroid: centroid and area of a polygon.
 * @customfunction POLYGON_CENTROIDAREA
 * @helpurl https://jpbardet.github.io/geoXL/index.html#POLYGON_CENTROIDAREA
 * @param {any[][]} polygon Vertices (two columns x, y); the first vertex may be repeated at the end
 * @returns {any[][]} Result (single value, row or table).
 */
function polygon_centroidArea(polygon) {
  return safe(() => G.polygon_area_centroid(pts2(polygon)));
}

/**
 * Weight and moment {W, M} of a polygonal block about point Pt. Inputs: XY vertices, g unit weight, Pt {x, y}. Per unit thickness.
 * @customfunction POLYGON_WEIGHTMOMENT
 * @helpurl https://jpbardet.github.io/geoXL/index.html#POLYGON_WEIGHTMOMENT
 * @param {any[][]} XY Block vertices (two columns x, y)
 * @param {any[][]} g Unit weight (kN/m³)
 * @param {any[][]} Pt Point {x, y} about which the moment is taken
 * @returns {any[][]} Result (single value, row or table).
 */
function polygon_weightMoment(XY, g, Pt) {
  return safe(() => G.polygon_weightMoment(pts2(XY), num(g), point(Pt)));
}

/**
 * Resultant force and moment {F, M} of a piecewise-linear pressure profile about depth d0. Inputs: d depth column, p pressure column, d0.
 * @customfunction PRESSURE_FORCEMOMENT
 * @helpurl https://jpbardet.github.io/geoXL/index.html#PRESSURE_FORCEMOMENT
 * @param {any[][]} d Column of depths (m)
 * @param {any[][]} p Column of pressures (kPa)
 * @param {any[][]} d0 Depth of the moment centre (m)
 * @returns {any[][]} Result (single value, row or table).
 */
function pressure_forceMoment(d, p, d0) {
  return safe(() => G.pressure_forceMoment(colNum(d), colNum(p), num(d0)));
}

/**
 * Rankine active coefficient for a sloping backfill. Inputs: phi, beta (deg).
 * @customfunction RANKINE_K
 * @helpurl https://jpbardet.github.io/geoXL/index.html#RANKINE_K
 * @param {any[][]} phi Friction angle φ (degrees)
 * @param {any[][]} beta Backfill slope β (degrees)
 * @returns {any[][]} Result (single value, row or table).
 */
function Rankine_K(phi, beta) {
  return safeLifted(() => lift((e_phi, e_beta) => G.Rankine_K(num(e_phi), num(e_beta)), [phi, beta], []));
}

/**
 * Rankine stresses on a wall from the Mohr circle (labelled table). Inputs: phi, coh, alpha wall angle, beta slope (deg), gz vertical stress, passive (omit = active).
 * @customfunction RANKINE_STRESS
 * @helpurl https://jpbardet.github.io/geoXL/index.html#RANKINE_STRESS
 * @param {any[][]} phi Friction angle φ (degrees)
 * @param {any[][]} coh Cohesion c (kPa)
 * @param {any[][]} alpha Wall angle α (degrees)
 * @param {any[][]} beta Backfill slope β (degrees)
 * @param {any[][]} gz Vertical stress γ·z (kPa)
 * @param {any[][]} [passive] Omit for active; any value for passive
 * @returns {any[][]} Result (single value, row or table).
 */
function Rankine_stress(phi, coh, alpha, beta, gz, passive) {
  return safe(() => G.Rankine_stress(num(phi), num(coh), num(alpha), num(beta), num(gz), opt(passive)));
}

/**
 * Sine of an angle in degrees.
 * @customfunction SIND
 * @helpurl https://jpbardet.github.io/geoXL/index.html#SIND
 * @param {any[][]} X Angle (degrees)
 * @returns {any[][]} Result (single value, row or table).
 */
function sind(X) {
  return safeLifted(() => lift((e_X) => G.sind(num(e_X)), [X], []));
}

/**
 * Allowable bearing pressure from SPT N averaged between D and D + B. Inputs: z, N columns (or scalar N), B, D (m), author: bowles, meyerhof, terzaghi.
 * @customfunction SPT_BEARING
 * @helpurl https://jpbardet.github.io/geoXL/index.html#SPT_BEARING
 * @param {any[][]} z Column of depths (m)
 * @param {any[][]} N Column of SPT N values (or a single N)
 * @param {any[][]} B Footing width (m)
 * @param {any[][]} D Footing depth (m)
 * @param {any[][]} author "bowles", "meyerhof" or "terzaghi"
 * @returns {any[][]} Result (single value, row or table).
 */
function SPT_bearing(z, N, B, D, author) {
  return safe(() => G.SPT_bearing(colNum(z), colNum(N), num(B), num(D), str(author)));
}

/**
 * Burland–Burbidge settlement from SPT. Inputs: z, N columns, q pressure, B width, optional LB = L/B, D_ depth, H layer thickness, sc preconsolidation pressure.
 * @customfunction SPT_BURLANDBURBIDGE_SETTLEMENT
 * @helpurl https://jpbardet.github.io/geoXL/index.html#SPT_BURLANDBURBIDGE_SETTLEMENT
 * @param {any[][]} z Column of depths (m)
 * @param {any[][]} N Column of SPT N values
 * @param {any[][]} q Applied pressure (kPa)
 * @param {any[][]} B Footing width (m)
 * @param {any[][]} [LB] Ratio L/B, optional (shape correction)
 * @param {any[][]} [D_] Footing depth (m), optional, default 0
 * @param {any[][]} [H] Thickness of the compressible layer (m), optional (thin-layer correction)
 * @param {any[][]} [sc] Preconsolidation pressure σ'p (kPa), optional
 * @returns {any[][]} Result (single value, row or table).
 */
function SPT_BurlandBurbidge_settlement(z, N, q, B, LB, D_, H, sc) {
  return safe(() => G.SPT_BurlandBurbidge_settlement(colNum(z), colNum(N), num(q), num(B), opt(LB), opt(D_), opt(H), opt(sc)));
}

/**
 * SPT overburden correction CN = 10/√σ'v (kPa), capped at 2.
 * @customfunction SPT_CORRECTION_DEPTH
 * @helpurl https://jpbardet.github.io/geoXL/index.html#SPT_CORRECTION_DEPTH
 * @param {any[][]} s Effective vertical stress σ'v (kPa)
 * @returns {any[][]} Result (single value, row or table).
 */
function SPT_correction_depth(s) {
  return safeLifted(() => lift((e_s) => G.SPT_correction_depth(num(e_s)), [s], []));
}

/**
 * Friction angle (deg) from SPT N. author: peck, schmertman (needs s and pa), hatanaka (needs s, in kPa).
 * @customfunction SPT_CORRELATION_FRICTION
 * @helpurl https://jpbardet.github.io/geoXL/index.html#SPT_CORRELATION_FRICTION
 * @param {any[][]} N SPT blow count N
 * @param {any[][]} author "peck", "schmertman" or "hatanaka"
 * @param {any[][]} [s] Effective vertical stress σ'v (needed by schmertman and hatanaka; kPa for hatanaka)
 * @param {any[][]} [pa] Atmospheric pressure in the same unit as s (needed by schmertman)
 * @returns {any[][]} Result (single value, row or table).
 */
function SPT_Correlation_friction(N, author, s, pa) {
  return safeLifted(() => lift((e_N, e_author, e_s, e_pa) => G.SPT_Correlation_friction(num(e_N), e_author, e_s, e_pa), [N, author, s, pa], []));
}

/**
 * Resultant force, moment and point of application of normal/shear stresses varying linearly along segment AB. Inputs: A, B {x, y}; sa, sb {normal, shear}.
 * @customfunction STRESS_FORCEMOMENT
 * @helpurl https://jpbardet.github.io/geoXL/index.html#STRESS_FORCEMOMENT
 * @param {any[][]} A Segment start {x, y}
 * @param {any[][]} B Segment end {x, y}
 * @param {any[][]} sa Stresses at A {normal, shear}
 * @param {any[][]} sb Stresses at B {normal, shear}
 * @returns {any[][]} Result (single value, row or table).
 */
function stress_ForceMoment(A, B, sa, sb) {
  return safe(() => G.stress_ForceMoment(point(A), point(B), point(sa), point(sb)));
}

/**
 * Normal and shear stress {s, t} on a plane at angle alpha (deg) using the Mohr circle pole. Inputs: sa, ta pole stresses, alpha, sm circle centre.
 * @customfunction STRESS_MOHR
 * @helpurl https://jpbardet.github.io/geoXL/index.html#STRESS_MOHR
 * @param {any[][]} sa Normal stress at the pole (kPa)
 * @param {any[][]} ta Shear stress at the pole (kPa)
 * @param {any[][]} alpha Plane angle (degrees)
 * @param {any[][]} sm Centre of the Mohr circle (kPa)
 * @returns {any[][]} Result (single value, row or table).
 */
function Stress_Mohr(sa, ta, alpha, sm) {
  return safeLifted(() => liftRows((e_sa, e_ta, e_alpha, e_sm) => G.Stress_Mohr(num(e_sa), num(e_ta), num(e_alpha), num(e_sm)), [sa, ta, alpha, sm], []));
}

/**
 * Mohr–Coulomb failure circle through a known stress (sa, ta) on plane A (deg): centre, radius, pole, principal and failure stresses. passive: omit = active.
 * @customfunction STRESS_MOHR_FAILURE
 * @helpurl https://jpbardet.github.io/geoXL/index.html#STRESS_MOHR_FAILURE
 * @param {any[][]} sa Known normal stress (kPa)
 * @param {any[][]} ta Known shear stress (kPa)
 * @param {any[][]} A Angle of the plane carrying (sa, ta) (degrees)
 * @param {any[][]} phi Friction angle φ (degrees)
 * @param {any[][]} coh Cohesion c (kPa)
 * @param {any[][]} [passive] Omit for active; any value for passive
 * @returns {any[][]} Result (single value, row or table).
 */
function stress_Mohr_failure(sa, ta, A, phi, coh, passive) {
  return safe(() => G.stress_Mohr_failure(num(sa), num(ta), num(A), num(phi), num(coh), opt(passive)));
}

/**
 * Mohr circle and pole from stresses on two perpendicular planes: centre, radius, pole, principal stresses. Inputs: sa, ta on plane A (deg), sb normal stress on the perpendicular plane.
 * @customfunction STRESS_MOHR_POLE
 * @helpurl https://jpbardet.github.io/geoXL/index.html#STRESS_MOHR_POLE
 * @param {any[][]} sa Normal stress on plane A (kPa)
 * @param {any[][]} ta Shear stress on plane A (kPa)
 * @param {any[][]} A Angle of plane A (degrees)
 * @param {any[][]} sb Normal stress on the perpendicular plane (kPa)
 * @returns {any[][]} Result (single value, row or table).
 */
function stress_Mohr_pole(sa, ta, A, sb) {
  return safe(() => G.stress_Mohr_pole(num(sa), num(ta), num(A), num(sb)));
}

/**
 * Plotting points {σ, τ} of a Mohr circle with centre sm and radius rValue (header + 21 points).
 * @customfunction STRESS_MOHRCIRCLE
 * @helpurl https://jpbardet.github.io/geoXL/index.html#STRESS_MOHRCIRCLE
 * @param {any[][]} sm Centre of the circle (kPa)
 * @param {any[][]} rValue Radius of the circle (kPa)
 * @returns {any[][]} Result (single value, row or table).
 */
function stress_Mohrcircle(sm, rValue) {
  return safe(() => G.stress_Mohrcircle(num(sm), num(rValue)));
}

/**
 * Tangent of an angle in degrees.
 * @customfunction TAND
 * @helpurl https://jpbardet.github.io/geoXL/index.html#TAND
 * @param {any[][]} X Angle (degrees)
 * @returns {any[][]} Result (single value, row or table).
 */
function tand(X) {
  return safeLifted(() => lift((e_X) => G.tand(num(e_X)), [X], []));
}

/**
 * Plotting coordinates of stress vectors along segment ptA–ptB (rows separated by #N/A). Inputs: sigA, sigB {normal, shear}, M divisions, s display scale.
 * @customfunction VECTOR_LINE
 * @helpurl https://jpbardet.github.io/geoXL/index.html#VECTOR_LINE
 * @param {any[][]} ptA Segment start {x, y}
 * @param {any[][]} ptB Segment end {x, y}
 * @param {any[][]} sigA Stresses at A {normal, shear}
 * @param {any[][]} sigB Stresses at B {normal, shear}
 * @param {any[][]} M Number of vectors drawn along the segment
 * @param {any[][]} s Display scale (length per unit stress)
 * @returns {any[][]} Result (single value, row or table).
 */
function vector_line(ptA, ptB, sigA, sigB, M, s) {
  return safe(() => G.vector_line(point(ptA), point(ptB), point(sigA), point(sigB), num(M), num(s)));
}

/**
 * Hydrostatic pore pressure at depth below the water table (0 above it). Optional unitWeight, default 9.81 kN/m³.
 * @customfunction WATER_PRESSURE
 * @helpurl https://jpbardet.github.io/geoXL/index.html#WATER_PRESSURE
 * @param {any[][]} depth Depth below the ground surface (m)
 * @param {any[][]} waterTableDepth Depth of the water table below the ground surface (m)
 * @param {any[][]} [unitWeight] Unit weight of water, optional, default 9.81 kN/m³
 * @returns {any[][]} Result (single value, row or table).
 */
function water_pressure(depth, waterTableDepth, unitWeight) {
  return safeLifted(() => lift((e_depth, e_waterTableDepth, e_unitWeight) => G.water_pressure(num(e_depth), num(e_waterTableDepth), e_unitWeight), [depth, waterTableDepth, unitWeight], []));
}

/**
 * Strip-load stress ratios in the order {sx, sy, sxy} (retaining-wall convention). Inputs: B full width, X offset, Y depth.
 * @customfunction ELASTIC_STRIPLOAD_XY
 * @helpurl https://jpbardet.github.io/geoXL/index.html#ELASTIC_STRIPLOAD_XY
 * @param {any[][]} B Full width of the strip (m)
 * @param {any[][]} X Horizontal offset from the strip centre (m)
 * @param {any[][]} Y Depth (m)
 * @returns {any[][]} Result (single value, row or table).
 */
function elastic_stripload_xy(B, X, Y) {
  return safeLifted(() => liftRows((e_B, e_X, e_Y) => G.elastic_stripload_xy(num(e_B), num(e_X), num(e_Y)), [B, X, Y], []));
}

/**
 * Plotting points of a circle or arc (header + n+1 points). Inputs: pt centre {x, y}, rValue radius, optional ta_, tb_ start/end angles (deg), n_ points (default 100).
 * @customfunction CIRCLE
 * @helpurl https://jpbardet.github.io/geoXL/index.html#CIRCLE
 * @param {any[][]} pt Centre {x, y}
 * @param {any[][]} rValue Radius
 * @param {any[][]} [ta_] Start angle (degrees), optional, default 0
 * @param {any[][]} [tb_] End angle (degrees), optional, default 360
 * @param {any[][]} [n_] Number of segments, optional, default 100
 * @returns {any[][]} Result (single value, row or table).
 */
function circle(pt, rValue, ta_, tb_, n_) {
  return safe(() => G.circle(point(pt), num(rValue), opt(ta_), opt(tb_), opt(n_)));
}

/**
 * LCPC CPT pile parameters {category, kc, β, qs max (MPa)}. Inputs: qc (MPa), soil (clay, silt, sand, gravel, chalk), pile (drilled, cast, steel, concrete).
 * @customfunction CLASSIFY_CPT_LCPC
 * @helpurl https://jpbardet.github.io/geoXL/index.html#CLASSIFY_CPT_LCPC
 * @param {any[][]} qc Cone resistance qc (MPa)
 * @param {any[][]} soil "clay", "silt", "sand", "gravel" or "chalk"
 * @param {any[][]} pile Pile type: "drilled", "cast", "steel" or "concrete"
 * @returns {any[][]} Result (single value, row or table).
 */
function classify_CPT_LCPC(qc, soil, pile) {
  return safeLifted(() => liftRows((e_qc, e_soil, e_pile) => G.classify_CPT_LCPC(num(e_qc), e_soil, e_pile), [qc, soil, pile], []));
}

/**
 * LCPC pressuremeter pile parameters {kp, shaft-friction curve number}. Inputs: pl limit pressure (MPa), soil (clay … rock), pile installation type.
 * @customfunction CLASSIFY_PMT_LCPC
 * @helpurl https://jpbardet.github.io/geoXL/index.html#CLASSIFY_PMT_LCPC
 * @param {any[][]} pl Limit pressure pl (MPa)
 * @param {any[][]} soil Soil type ("clay" … "rock")
 * @param {any[][]} pile Pile installation type
 * @returns {any[][]} Result (single value, row or table).
 */
function classify_PMT_LCPC(pl, soil, pile) {
  return safeLifted(() => liftRows((e_pl, e_soil, e_pile) => G.classify_PMT_LCPC(num(e_pl), e_soil, e_pile), [pl, soil, pile], []));
}

/**
 * Average CPT qc below a footing: D to D + 1.5B (meyerhof) or D to D + B (robertson). Inputs: z, qc columns, D, B, author.
 * @customfunction CPT_AVERAGE_QC
 * @helpurl https://jpbardet.github.io/geoXL/index.html#CPT_AVERAGE_QC
 * @param {any[][]} z Column of depths (m)
 * @param {any[][]} qc Column of cone resistances qc
 * @param {any[][]} D Footing depth (m)
 * @param {any[][]} B Footing width (m)
 * @param {any[][]} author "meyerhof" (D to D + 1.5B) or "robertson" (D to D + B)
 * @returns {any[][]} Result (single value, row or table).
 */
function CPT_average_qc(z, qc, D, B, author) {
  return safe(() => G.CPT_average_qc(colNum(z), colNum(qc), num(D), num(B), str(author)));
}

/**
 * Meyerhof CPT bearing factor with water-table corrections Cw1, Cw2. Inputs: B, D, dw water depth, optional m (any value selects the 12.19 divisor).
 * @customfunction CPT_BEARING_MEYERHOF_K
 * @helpurl https://jpbardet.github.io/geoXL/index.html#CPT_BEARING_MEYERHOF_K
 * @param {any[][]} B Footing width (m)
 * @param {any[][]} D Footing depth (m)
 * @param {any[][]} dw Depth of the water table (m)
 * @param {any[][]} [m] Optional: any value selects the 12.19 divisor instead of 40
 * @returns {any[][]} Result (single value, row or table).
 */
function CPT_bearing_Meyerhof_k(B, D, dw, m) {
  return safeLifted(() => lift((e_B, e_D, e_dw, e_m) => G.CPT_bearing_Meyerhof_k(num(e_B), num(e_D), num(e_dw), e_m), [B, D, dw, m], []));
}

/**
 * Robertson CPT bearing factor k from chart fits. Inputs: B, D, shape (square, circular), density (dense, loose).
 * @customfunction CPT_BEARING_ROBERTSON_K
 * @helpurl https://jpbardet.github.io/geoXL/index.html#CPT_BEARING_ROBERTSON_K
 * @param {any[][]} B Footing width (m)
 * @param {any[][]} D Footing depth (m)
 * @param {any[][]} shape "square" or "circular"
 * @param {any[][]} density "dense" or "loose"
 * @returns {any[][]} Result (single value, row or table).
 */
function CPT_bearing_Robertson_k(B, D, shape, density) {
  return safeLifted(() => lift((e_B, e_D, e_shape, e_density) => G.CPT_bearing_Robertson_k(num(e_B), num(e_D), e_shape, e_density), [B, D, shape, density], []));
}

/**
 * Friction angle (deg) from CPT qc. Inputs: z depth, qc, g unit weight (consistent units), author_: robertson, kulhawy, ricceri, lee.
 * @customfunction CPT_FRICTION_ANGLE
 * @helpurl https://jpbardet.github.io/geoXL/index.html#CPT_FRICTION_ANGLE
 * @param {any[][]} z Depth (m)
 * @param {any[][]} qc Cone resistance qc (same unit as γ·z, so qc/σv is dimensionless)
 * @param {any[][]} g Unit weight γ (consistent with qc and z)
 * @param {any[][]} author_ "robertson", "kulhawy", "ricceri" or "lee"
 * @returns {any[][]} Result (single value, row or table).
 */
function CPT_friction_angle(z, qc, g, author_) {
  return safeLifted(() => lift((e_z, e_qc, e_g, e_author_) => G.CPT_friction_angle(num(e_z), num(e_qc), num(e_g), e_author_), [z, qc, g, author_], []));
}

/**
 * Layer thicknesses from an Olsen SCN profile (header + thickness, SCN+3). Input: cpt table with depth in column 1 and SCN in column 6.
 * @customfunction CPT_OLSEN_LAYERS
 * @helpurl https://jpbardet.github.io/geoXL/index.html#CPT_OLSEN_LAYERS
 * @param {any[][]} cpt Output table of CPT_Olsen_SCN with depth in column 1 and SCN in column 6
 * @returns {any[][]} Result (single value, row or table).
 */
function CPT_Olsen_layers(cpt) {
  return safe(() => G.CPT_Olsen_layers(mat(cpt)));
}

/**
 * Olsen soil classification number (SCN) for a CPT profile: {Qn, Fr, SCN, name}. Inputs: dqs (depth, qc, fs), pa, n exponent, g, gw, dw.
 * @customfunction CPT_OLSEN_SCN
 * @helpurl https://jpbardet.github.io/geoXL/index.html#CPT_OLSEN_SCN
 * @param {any[][]} dqs Three-column table {depth, qc, fs}
 * @param {any[][]} pa Atmospheric pressure (same unit as qc)
 * @param {any[][]} n Stress exponent for normalization
 * @param {any[][]} g Total unit weight of the soil
 * @param {any[][]} gw Unit weight of water
 * @param {any[][]} dw Depth of the water table
 * @returns {any[][]} Result (single value, row or table).
 */
function CPT_Olsen_SCN(dqs, pa, n, g, gw, dw) {
  return safe(() => G.CPT_Olsen_SCN(mat(dqs), num(pa), num(n), num(g), num(gw), num(dw)));
}

/**
 * Helper for CPT_Olsen_SCN: zone index (1–7) for one point. Inputs: Qt normalized tip resistance, Fr friction ratio (%).
 * @customfunction CPT_OLSEN_SCN_ROW
 * @helpurl https://jpbardet.github.io/geoXL/index.html#CPT_OLSEN_SCN_ROW
 * @param {any[][]} Qt Normalized tip resistance
 * @param {any[][]} Fr Friction ratio (%)
 * @returns {any[][]} Result (single value, row or table).
 */
function CPT_Olsen_SCN_row(Qt, Fr) {
  return safeLifted(() => lift((e_Qt, e_Fr) => G.CPT_Olsen_SCN_row(num(e_Qt), num(e_Fr)), [Qt, Fr], []));
}

/**
 * Boundary curve A–H of the Robertson CPT classification chart: Qt at friction ratio x_ (%). sel: letter of the curve.
 * @customfunction CPT_ROBERTSON_CLASSIFICATION_CHART
 * @helpurl https://jpbardet.github.io/geoXL/index.html#CPT_ROBERTSON_CLASSIFICATION_CHART
 * @param {any[][]} x_ Friction ratio Fr (%)
 * @param {any[][]} sel Boundary curve letter "A" … "H"
 * @returns {any[][]} Result (single value, row or table).
 */
function CPT_Robertson_classification_chart(x_, sel) {
  return safeLifted(() => lift((e_x_, e_sel) => G.CPT_Robertson_classification_chart(num(e_x_), e_sel), [x_, sel], []));
}

/**
 * Layer thicknesses from Robertson SBT and SBT(Ic) profiles. Input: cpt table with depth in column 1, SBT in column 6, SBT(Ic) in column 9.
 * @customfunction CPT_ROBERTSON_LAYERS
 * @helpurl https://jpbardet.github.io/geoXL/index.html#CPT_ROBERTSON_LAYERS
 * @param {any[][]} cpt Output table of CPT_Robertson_SBT (depth in column 1, SBT in column 6, SBT(Ic) in column 9)
 * @returns {any[][]} Result (single value, row or table).
 */
function CPT_Robertson_layers(cpt) {
  return safe(() => G.CPT_Robertson_layers(mat(cpt)));
}

/**
 * Robertson soil behaviour type for a CPT profile: {Qn, Fr, SBT, name, Ic, SBT(Ic), name}. Inputs: dqs (depth, qc, fs), pa, n exponent, g, gw, dw.
 * @customfunction CPT_ROBERTSON_SBT
 * @helpurl https://jpbardet.github.io/geoXL/index.html#CPT_ROBERTSON_SBT
 * @param {any[][]} dqs Three-column table {depth, qc, fs}
 * @param {any[][]} pa Atmospheric pressure (same unit as qc)
 * @param {any[][]} n Stress exponent for normalization
 * @param {any[][]} g Total unit weight of the soil
 * @param {any[][]} gw Unit weight of water
 * @param {any[][]} dw Depth of the water table
 * @returns {any[][]} Result (single value, row or table).
 */
function CPT_Robertson_SBT(dqs, pa, n, g, gw, dw) {
  return safe(() => G.CPT_Robertson_SBT(mat(dqs), num(pa), num(n), num(g), num(gw), num(dw)));
}

/**
 * Helper for CPT_Robertson_SBT: SBT zone (1–9, 0 = outside chart) for one point. Inputs: Qt_, Fr_ (%).
 * @customfunction CPT_ROBERTSON_SBT_ROW
 * @helpurl https://jpbardet.github.io/geoXL/index.html#CPT_ROBERTSON_SBT_ROW
 * @param {any[][]} [Qt_] Normalized tip resistance
 * @param {any[][]} [Fr_] Friction ratio (%)
 * @returns {any[][]} Result (single value, row or table).
 */
function CPT_Robertson_SBT_row(Qt_, Fr_) {
  return safeLifted(() => lift((e_Qt_, e_Fr_) => G.CPT_Robertson_SBT_row(e_Qt_, e_Fr_), [Qt_, Fr_], []));
}

/**
 * Schmertmann settlement from a CPT profile (total settlement + layer table). Inputs: zqc (depth, qc), k = E/qc, B, D, L, qn net pressure, g unit weight.
 * @customfunction CPT_SETTLEMENT
 * @helpurl https://jpbardet.github.io/geoXL/index.html#CPT_SETTLEMENT
 * @param {any[][]} zqc Two-column table {depth (m), qc}
 * @param {any[][]} k Modulus factor E/qc (e.g. 2.5 square, 3.5 strip)
 * @param {any[][]} B Footing width (m)
 * @param {any[][]} D Footing depth (m)
 * @param {any[][]} L Footing length (m)
 * @param {any[][]} qn Net applied pressure (same unit as qc)
 * @param {any[][]} g Unit weight above footing level
 * @returns {any[][]} Result (single value, row or table).
 */
function CPT_settlement(zqc, k, B, D, L, qn, g) {
  return safe(() => G.CPT_settlement(mat(zqc), num(k), num(B), num(D), num(L), num(qn), num(g)));
}

/**
 * Schmertmann strain influence factor Iz at depth z (accepts a column). Inputs: z, B, D, L, qn net pressure, g unit weight.
 * @customfunction CPT_SETTLEMENT_IZ
 * @helpurl https://jpbardet.github.io/geoXL/index.html#CPT_SETTLEMENT_IZ
 * @param {any[][]} z Depth(s) (m)
 * @param {any[][]} B Footing width (m)
 * @param {any[][]} D Footing depth (m)
 * @param {any[][]} L Footing length (m)
 * @param {any[][]} qn Net applied pressure
 * @param {any[][]} g Unit weight above footing level
 * @returns {any[][]} Result (single value, row or table).
 */
function CPT_settlement_Iz(z, B, D, L, qn, g) {
  return safeLifted(() => lift((e_z, e_B, e_D, e_L, e_qn, e_g) => G.CPT_settlement_Iz(num(e_z), num(e_B), num(e_D), num(e_L), num(e_qn), num(e_g)), [z, B, D, L, qn, g], []));
}

/**
 * Perpendicular distances from point M {x, y} to the line through each segment of polyline pl.
 * @customfunction DISTANCE_POINT_POLYLINE
 * @helpurl https://jpbardet.github.io/geoXL/index.html#DISTANCE_POINT_POLYLINE
 * @param {any[][]} M Point {x, y}
 * @param {any[][]} pl Polyline vertices (two columns x, y)
 * @returns {any[][]} Result (single value, row or table).
 */
function distance_point_polyline(M, pl) {
  return safe(() => G.distance_point_polyline(point(M), pts2(pl)));
}

/**
 * Bilinear interpolation in a 2-D table. Inputs: x, Y query; XY table with x values in row 1 and y values in column 1 (ascending).
 * @customfunction INTERPOLATE_TABLE
 * @helpurl https://jpbardet.github.io/geoXL/index.html#INTERPOLATE_TABLE
 * @param {any[][]} x Value along the first row of the table
 * @param {any[][]} Y Value along the first column of the table
 * @param {any[][]} XY Table: x values in row 1, y values in column 1 (ascending), data inside
 * @returns {any[][]} Result (single value, row or table).
 */
function interpolate_table(x, Y, XY) {
  const XY_c = mat(XY);
  return safeLifted(() => lift((e_x, e_Y, e_XY) => G.interpolate_table(num(e_x), num(e_Y), e_XY), [x, Y, XY_c], [2]));
}

/**
 * Intersection points {x, y} of a polyline with a circle. Inputs: pline vertices, center {x, y}, radius.
 * @customfunction INTERSECTION_POLYLINE_CIRCLE
 * @helpurl https://jpbardet.github.io/geoXL/index.html#INTERSECTION_POLYLINE_CIRCLE
 * @param {any[][]} pline Polyline vertices (two columns x, y)
 * @param {any[][]} center Circle centre {x, y}
 * @param {any[][]} radius Circle radius
 * @returns {any[][]} Result (single value, row or table).
 */
function intersection_polyline_circle(pline, center, radius) {
  return safe(() => G.intersection_polyline_circle(pts2(pline), point(center), num(radius)));
}

/**
 * Intersections of polyline pl with polyline ql: {parameter, x, y}. Optional m keeps the m points with the largest x.
 * @customfunction INTERSECTION_POLYLINE_POLYLINE
 * @helpurl https://jpbardet.github.io/geoXL/index.html#INTERSECTION_POLYLINE_POLYLINE
 * @param {any[][]} pl First polyline (two columns x, y)
 * @param {any[][]} ql Second polyline (two columns x, y)
 * @param {any[][]} [m] Optional: keep only the m intersections with the largest x
 * @returns {any[][]} Result (single value, row or table).
 */
function intersection_polyline_polyline(pl, ql, m) {
  return safe(() => G.intersection_polyline_polyline(pts2(pl), pts2(ql), opt(m)));
}

/**
 * Helper for intersection_polyline_polyline: segment parameter of the intersection between one segment and a polyline. Input: flattened row (segment, then polyline).
 * @customfunction INTERSECTION_SEGMENT_POLYLINE
 * @helpurl https://jpbardet.github.io/geoXL/index.html#INTERSECTION_SEGMENT_POLYLINE
 * @param {any[][]} row Flattened row: segment end points, then the polyline coordinates
 * @returns {any[][]} Result (single value, row or table).
 */
function intersection_segment_polyline(row) {
  return safe(() => G.intersection_segment_polyline(colNum(row)));
}

/**
 * Intersection {x, y} of a log spiral (centre pt, radius rValue, angle phi) with segment AB_ by Newton iteration.
 * @customfunction INTERSECTION_SPIRAL_SEGMENT
 * @helpurl https://jpbardet.github.io/geoXL/index.html#INTERSECTION_SPIRAL_SEGMENT
 * @param {any[][]} AB_ Segment end points (2 rows × 2 columns)
 * @param {any[][]} pt Spiral centre {x, y}
 * @param {any[][]} rValue Spiral initial radius
 * @param {any[][]} phi Friction angle φ of the spiral (degrees)
 * @param {any[][]} [x_] Optional starting guess
 * @param {any[][]} [iter_] Leave empty (internal counter)
 * @returns {any[][]} Result (single value, row or table).
 */
function intersection_spiral_segment(AB_, pt, rValue, phi, x_, iter_) {
  return safe(() => G.intersection_spiral_segment(pts2(AB_), point(pt), num(rValue), num(phi), opt(x_), opt(iter_)));
}

/**
 * Broms lateral capacity of a short pile in cohesionless soil: {V; depth of max moment; Mmax}. Inputs: D embedment, phi, g, B diameter, e load eccentricity.
 * @customfunction LATERALPILE_BROMS_COHESIONLESS
 * @helpurl https://jpbardet.github.io/geoXL/index.html#LATERALPILE_BROMS_COHESIONLESS
 * @param {any[][]} D Embedded length of the pile (m)
 * @param {any[][]} phi Friction angle φ (degrees)
 * @param {any[][]} g Unit weight γ (kN/m³)
 * @param {any[][]} B Pile diameter (m)
 * @param {any[][]} e Height of the load above the ground (m)
 * @returns {any[][]} Result (single value, row or table).
 */
function lateralpile_Broms_cohesionless(D, phi, g, B, e) {
  return safe(() => G.lateralpile_Broms_cohesionless(num(D), num(phi), num(g), num(B), num(e)));
}

/**
 * Broms short pile in cohesive soil: {required D; rotation depth Dr; depth of max moment; Mmax}. Inputs: V load, e eccentricity, Su, B diameter.
 * @customfunction LATERALPILE_BROMS_COHESIVE
 * @helpurl https://jpbardet.github.io/geoXL/index.html#LATERALPILE_BROMS_COHESIVE
 * @param {any[][]} V Lateral load (kN)
 * @param {any[][]} e Height of the load above the ground (m)
 * @param {any[][]} Su Undrained shear strength (kPa)
 * @param {any[][]} B Pile diameter (m)
 * @returns {any[][]} Result (single value, row or table).
 */
function lateralpile_Broms_cohesive(V, e, Su, B) {
  return safe(() => G.lateralpile_Broms_cohesive(num(V), num(e), num(Su), num(B)));
}

/**
 * Reaction, shear and moment diagrams along a Broms short pile in cohesionless soil (header + 100 points). Inputs: D, phi, g, B, e.
 * @customfunction LATERALPILE_COHESIONLESS_REACTION_SHEAR_MOMENT
 * @helpurl https://jpbardet.github.io/geoXL/index.html#LATERALPILE_COHESIONLESS_REACTION_SHEAR_MOMENT
 * @param {any[][]} D Embedded length (m)
 * @param {any[][]} phi Friction angle φ (degrees)
 * @param {any[][]} g Unit weight γ (kN/m³)
 * @param {any[][]} B Pile diameter (m)
 * @param {any[][]} e Height of the load above the ground (m)
 * @returns {any[][]} Result (single value, row or table).
 */
function lateralpile_cohesionless_reaction_shear_moment(D, phi, g, B, e) {
  return safe(() => G.lateralpile_cohesionless_reaction_shear_moment(num(D), num(phi), num(g), num(B), num(e)));
}

/**
 * Reaction, shear and moment diagrams along a Broms short pile in cohesive soil. Inputs: V, e, Su, B, optional D_ and Dr_ (default from Broms).
 * @customfunction LATERALPILE_COHESIVE_REACTION_SHEAR_MOMENT
 * @helpurl https://jpbardet.github.io/geoXL/index.html#LATERALPILE_COHESIVE_REACTION_SHEAR_MOMENT
 * @param {any[][]} V Lateral load (kN)
 * @param {any[][]} e Height of the load above the ground (m)
 * @param {any[][]} Su Undrained shear strength (kPa)
 * @param {any[][]} B Pile diameter (m)
 * @param {any[][]} [D_] Embedded length (m), optional (default from Broms)
 * @param {any[][]} [Dr_] Rotation depth (m), optional (default from Broms)
 * @returns {any[][]} Result (single value, row or table).
 */
function lateralpile_cohesive_reaction_shear_moment(V, e, Su, B, D_, Dr_) {
  return safe(() => G.lateralpile_cohesive_reaction_shear_moment(num(V), num(e), num(Su), num(B), opt(D_), opt(Dr_)));
}

/**
 * Plotting points of a log spiral r = rValue·exp(−θ·tanφ). Inputs: pt centre, rValue, phi (deg), optional ta_, tb_ angles (deg), n_ points (default 100).
 * @customfunction LOGSPIRAL
 * @helpurl https://jpbardet.github.io/geoXL/index.html#LOGSPIRAL
 * @param {any[][]} pt Centre {x, y}
 * @param {any[][]} rValue Initial radius r0
 * @param {any[][]} phi Friction angle φ (degrees)
 * @param {any[][]} [ta_] Start angle (degrees), optional, default 0
 * @param {any[][]} [tb_] End angle (degrees), optional, default −180
 * @param {any[][]} [n_] Number of segments, optional, default 100
 * @returns {any[][]} Result (single value, row or table).
 */
function logSpiral(pt, rValue, phi, ta_, tb_, n_) {
  return safe(() => G.logSpiral(point(pt), num(rValue), num(phi), opt(ta_), opt(tb_), opt(n_)));
}

/**
 * Schmertmann CPT pile capacity: ultimate, shaft, tip, average fs and qc. Inputs: table (z, qc, fs, optional soil), A tip area, p perimeter, L length, pile (steel, concrete, timber).
 * @customfunction PILE_CAPACITY_CPT_SCHMERTMAN
 * @helpurl https://jpbardet.github.io/geoXL/index.html#PILE_CAPACITY_CPT_SCHMERTMAN
 * @param {any[][]} z_qc_fs_soil Table {depth (m), qc, fs, optional soil ("clay", "silt", "sand", "gravel"; default sand)}
 * @param {any[][]} A Pile tip area (m²)
 * @param {any[][]} p Pile perimeter (m)
 * @param {any[][]} L Pile length (m)
 * @param {any[][]} pile "steel", "concrete" or "timber"
 * @returns {any[][]} Result (single value, row or table).
 */
function pile_capacity_CPT_Schmertman(z_qc_fs_soil, A, p, L, pile) {
  return safe(() => G.pile_capacity_CPT_Schmertman(mat(z_qc_fs_soil), num(A), num(p), num(L), str(pile)));
}

/**
 * Shaft resistance by the α method. Inputs: z, Su columns, pa atmospheric pressure, p perimeter, L length, author (terzaghi, reese).
 * @customfunction PILE_SHAFT_ALPHA
 * @helpurl https://jpbardet.github.io/geoXL/index.html#PILE_SHAFT_ALPHA
 * @param {any[][]} z Column of depths (m)
 * @param {any[][]} Su Column of undrained shear strengths (kPa)
 * @param {any[][]} pa Atmospheric pressure (kPa, about 101)
 * @param {any[][]} p Pile perimeter (m)
 * @param {any[][]} L Pile length (m)
 * @param {any[][]} author "terzaghi" or "reese"
 * @returns {any[][]} Result (single value, row or table).
 */
function pile_shaft_alpha(z, Su, pa, p, L, author) {
  return safe(() => G.pile_shaft_alpha(colNum(z), colNum(Su), num(pa), num(p), num(L), str(author)));
}

/**
 * Adhesion factor α. Inputs: Su = Su/pa (normalized), author (terzaghi, reese).
 * @customfunction PILE_SIDEFRICTION_ALPHA
 * @helpurl https://jpbardet.github.io/geoXL/index.html#PILE_SIDEFRICTION_ALPHA
 * @param {any[][]} Su Normalized strength Su/pa
 * @param {any[][]} author "terzaghi" or "reese"
 * @returns {any[][]} Result (single value, row or table).
 */
function pile_sidefriction_alpha(Su, author) {
  return safeLifted(() => lift((e_Su, e_author) => G.pile_sidefriction_alpha(num(e_Su), e_author), [Su, author], []));
}

/**
 * λ coefficient for pile shaft friction (Vijayvergiya–Focht chart fit). Input: D_ pile length in m (or ft if ft is supplied).
 * @customfunction PILE_SIDEFRICTION_LAMBDA
 * @helpurl https://jpbardet.github.io/geoXL/index.html#PILE_SIDEFRICTION_LAMBDA
 * @param {any[][]} D_ Pile length (m, or ft if ft is supplied)
 * @param {any[][]} [ft] Optional: any value means D_ is in feet
 * @returns {any[][]} Result (single value, row or table).
 */
function pile_sidefriction_lambda(D_, ft) {
  return safeLifted(() => lift((e_D_, e_ft) => G.pile_sidefriction_lambda(num(e_D_), e_ft), [D_, ft], []));
}

/**
 * Meyerhof ultimate tip pressure. Inputs: pa atmospheric pressure, cValue cohesion, phi (deg).
 * @customfunction PILE_TIP_MEYERHOF
 * @helpurl https://jpbardet.github.io/geoXL/index.html#PILE_TIP_MEYERHOF
 * @param {any[][]} pa Atmospheric pressure (kPa)
 * @param {any[][]} cValue Cohesion c (kPa)
 * @param {any[][]} phi Friction angle φ (degrees)
 * @returns {any[][]} Result (single value, row or table).
 */
function pile_tip_Meyerhof(pa, cValue, phi) {
  return safe(() => G.pile_tip_Meyerhof(num(pa), num(cValue), num(phi)));
}

/**
 * Meyerhof pile tip factors {Nq; Nc} for phi (deg).
 * @customfunction PILE_TIP_MEYERHOF_FACTOR
 * @helpurl https://jpbardet.github.io/geoXL/index.html#PILE_TIP_MEYERHOF_FACTOR
 * @param {any[][]} phi Friction angle φ (degrees)
 * @returns {any[][]} Result (single value, row or table).
 */
function pile_tip_meyerhof_factor(phi) {
  return safe(() => G.pile_tip_meyerhof_factor(num(phi)));
}

/**
 * Vesic ultimate tip pressure with reduced rigidity index: {qu; Nq; Nc; Irr; σm}. Inputs: sv vertical stress, pa, cValue cohesion, optional phi_ (deg), Dr_ relative density.
 * @customfunction PILE_TIP_VESIC
 * @helpurl https://jpbardet.github.io/geoXL/index.html#PILE_TIP_VESIC
 * @param {any[][]} sv Vertical effective stress at the tip (kPa)
 * @param {any[][]} pa Atmospheric pressure (kPa)
 * @param {any[][]} cValue Cohesion c (kPa)
 * @param {any[][]} [phi_] Friction angle φ (degrees), optional, default 0
 * @param {any[][]} [Dr_] Relative density (fraction), optional, default 0.1
 * @returns {any[][]} Result (single value, row or table).
 */
function pile_tip_vesic(sv, pa, cValue, phi_, Dr_) {
  return safe(() => G.pile_tip_vesic(num(sv), num(pa), num(cValue), opt(phi_), opt(Dr_)));
}

/**
 * Vesic pile tip factors {Nq; Nc}. Inputs: phi (deg), Ir rigidity index.
 * @customfunction PILE_TIP_VESIC_FACTOR
 * @helpurl https://jpbardet.github.io/geoXL/index.html#PILE_TIP_VESIC_FACTOR
 * @param {any[][]} phi Friction angle φ (degrees)
 * @param {any[][]} Ir Rigidity index Ir
 * @returns {any[][]} Result (single value, row or table).
 */
function pile_tip_vesic_factor(phi, Ir) {
  return safe(() => G.pile_tip_vesic_factor(num(phi), num(Ir)));
}

/**
 * Ménard rheological factor α. Inputs: soil (peat, clay, silt, sand, gravel, rock), Epl = E/pl ratio.
 * @customfunction PMT_ALPHA
 * @helpurl https://jpbardet.github.io/geoXL/index.html#PMT_ALPHA
 * @param {any[][]} soil "peat", "clay", "silt", "sand", "gravel" or "rock"
 * @param {any[][]} Epl Ratio E/pl
 * @returns {any[][]} Result (single value, row or table).
 */
function PMT_alpha(soil, Epl) {
  return safeLifted(() => lift((e_soil, e_Epl) => G.PMT_alpha(e_soil, num(e_Epl)), [soil, Epl], []));
}

/**
 * Equivalent limit pressure ple (geometric mean from D to D + 1.5B) and equivalent embedment De: {ple; De}. Inputs: z, pl columns, B, D.
 * @customfunction PMT_BEARING_LIMITPRESSURE
 * @helpurl https://jpbardet.github.io/geoXL/index.html#PMT_BEARING_LIMITPRESSURE
 * @param {any[][]} z Column of test depths (m)
 * @param {any[][]} pl Column of limit pressures pl
 * @param {any[][]} B Footing width (m)
 * @param {any[][]} D Footing depth (m)
 * @returns {any[][]} Result (single value, row or table).
 */
function PMT_bearing_limitPressure(z, pl, B, D) {
  return safe(() => G.PMT_bearing_limitPressure(colNum(z), colNum(pl), num(B), num(D)));
}

/**
 * Ménard ultimate bearing pressure from pressuremeter data. Inputs: z, pl columns, B, D, L, soil category (I, II, III, IIIA), g unit weight.
 * @customfunction PMT_BEARING_MENARD
 * @helpurl https://jpbardet.github.io/geoXL/index.html#PMT_BEARING_MENARD
 * @param {any[][]} z Column of test depths (m)
 * @param {any[][]} pl Column of limit pressures pl
 * @param {any[][]} B Footing width (m)
 * @param {any[][]} D Footing depth (m)
 * @param {any[][]} L Footing length (m)
 * @param {any[][]} soil Ménard category "I", "II", "III" or "IIIA"
 * @param {any[][]} g Unit weight above footing level
 * @returns {any[][]} Result (single value, row or table).
 */
function PMT_bearing_Menard(z, pl, B, D, L, soil, g) {
  return safe(() => G.PMT_bearing_Menard(colNum(z), colNum(pl), num(B), num(D), num(L), str(soil), num(g)));
}

/**
 * Ménard bearing factor k, interpolated between strip and square. Inputs: H = De/(B/2), BL = B/L, soil category (I, II, III, IIIA).
 * @customfunction PMT_BEARING_MENARD_K
 * @helpurl https://jpbardet.github.io/geoXL/index.html#PMT_BEARING_MENARD_K
 * @param {any[][]} H Ratio De/(B/2)
 * @param {any[][]} BL Shape ratio B/L
 * @param {any[][]} soil Ménard category "I", "II", "III" or "IIIA"
 * @returns {any[][]} Result (single value, row or table).
 */
function PMT_bearing_Menard_k(H, BL, soil) {
  return safeLifted(() => lift((e_H, e_BL, e_soil) => G.PMT_bearing_Menard_k(num(e_H), num(e_BL), e_soil), [H, BL, soil], []));
}

/**
 * Plotting lines of the 16 Ménard sub-layers below a footing. Inputs: B, D.
 * @customfunction PMT_LAYERS
 * @helpurl https://jpbardet.github.io/geoXL/index.html#PMT_LAYERS
 * @param {any[][]} B Footing width (m)
 * @param {any[][]} D Footing depth (m)
 * @returns {any[][]} Result (single value, row or table).
 */
function PMT_layers(B, D) {
  return safe(() => G.PMT_layers(num(B), num(D)));
}

/**
 * Step-profile plotting coordinates from depth/value columns. Inputs: x depths, y values.
 * @customfunction STEPPROFILE
 * @helpurl https://jpbardet.github.io/geoXL/index.html#STEPPROFILE
 * @param {any[][]} x Column of depths
 * @param {any[][]} y Column of values
 * @returns {any[][]} Result (single value, row or table).
 */
function stepProfile(x, y) {
  return safe(() => G.stepProfile(colNum(x), colNum(y)));
}

/**
 * Alias of stepProfile for SPT profiles.
 * @customfunction SPT_STEPPROFILE
 * @helpurl https://jpbardet.github.io/geoXL/index.html#SPT_STEPPROFILE
 * @param {any[][]} x Column of depths
 * @param {any[][]} y Column of N values
 * @returns {any[][]} Result (single value, row or table).
 */
function SPT_stepProfile(x, y) {
  return safe(() => G.stepProfile(colNum(x), colNum(y)));
}

/**
 * Alias of stepProfile for CPT profiles.
 * @customfunction CPT_STEPPROFILE
 * @helpurl https://jpbardet.github.io/geoXL/index.html#CPT_STEPPROFILE
 * @param {any[][]} x Column of depths
 * @param {any[][]} y Column of CPT values
 * @returns {any[][]} Result (single value, row or table).
 */
function CPT_stepProfile(x, y) {
  return safe(() => G.stepProfile(colNum(x), colNum(y)));
}

/**
 * Alias of stepProfile for pressuremeter profiles.
 * @customfunction PMT_STAIRCASE
 * @helpurl https://jpbardet.github.io/geoXL/index.html#PMT_STAIRCASE
 * @param {any[][]} x Column of depths
 * @param {any[][]} y Column of pressuremeter values
 * @returns {any[][]} Result (single value, row or table).
 */
function PMT_staircase(x, y) {
  return safe(() => G.stepProfile(colNum(x), colNum(y)));
}

/**
 * Centroid and area of a polygon: {"Centroid =", xc, yc; "Area =", A, ""}. Input: vertices in two columns (first vertex may be repeated).
 * @customfunction POLYGON_AREA_CENTROID
 * @helpurl https://jpbardet.github.io/geoXL/index.html#POLYGON_AREA_CENTROID
 * @param {any[][]} polygon Vertices (two columns x, y); the first vertex may be repeated at the end
 * @returns {any[][]} Result (single value, row or table).
 */
function polygon_area_centroid(polygon) {
  return safe(() => G.polygon_area_centroid(pts2(polygon)));
}

/**
 * TRUE if point (x, y) lies below exactly one edge of the polygon (xa, ya) — vertical ray test.
 * @customfunction POLYGON_INSIDE
 * @helpurl https://jpbardet.github.io/geoXL/index.html#POLYGON_INSIDE
 * @param {any[][]} x x of the point
 * @param {any[][]} y y of the point
 * @param {any[][]} xa Column of polygon x coordinates
 * @param {any[][]} ya Column of polygon y coordinates
 * @returns {any[][]} Result (single value, row or table).
 */
function Polygon_inside(x, y, xa, ya) {
  const xa_c = colNum(xa);
  const ya_c = colNum(ya);
  return safeLifted(() => lift((e_x, e_y, e_xa, e_ya) => G.Polygon_inside(num(e_x), num(e_y), e_xa, e_ya), [x, y, xa_c, ya_c], [2, 3]));
}

/**
 * Linear interpolation of y at X along a polyline (two columns), searching for the segment starting at or below X. Used for root finding.
 * @customfunction POLYLINE_SUBDIVIDE
 * @helpurl https://jpbardet.github.io/geoXL/index.html#POLYLINE_SUBDIVIDE
 * @param {any[][]} X x value
 * @param {any[][]} polyline Polyline vertices (two columns x, y)
 * @returns {any[][]} Result (single value, row or table).
 */
function polyline_subdivide(X, polyline) {
  const polyline_c = pts2(polyline);
  return safeLifted(() => lift((e_X, e_polyline) => G.polyline_subdivide(num(e_X), e_polyline), [X, polyline_c], [1]));
}

/**
 * Schmertmann strain influence factor at depth z. Inputs: Imax, B, za footing depth, optional BL_ (0 = square, 1 = strip).
 * @customfunction SETTLEMENT_SCHMERTMAN_FACTOR
 * @helpurl https://jpbardet.github.io/geoXL/index.html#SETTLEMENT_SCHMERTMAN_FACTOR
 * @param {any[][]} z Depth (m)
 * @param {any[][]} Imax Peak strain influence factor Izp
 * @param {any[][]} B Footing width (m)
 * @param {any[][]} za Depth of the footing base (m)
 * @param {any[][]} [BL_] Optional shape: 0 = square/axisymmetric (default), 1 = strip
 * @returns {any[][]} Result (single value, row or table).
 */
function settlement_Schmertman_factor(z, Imax, B, za, BL_) {
  return safeLifted(() => lift((e_z, e_Imax, e_B, e_za, e_BL_) => G.settlement_Schmertman_factor(num(e_z), num(e_Imax), num(e_B), num(e_za), e_BL_), [z, Imax, B, za, BL_], []));
}

/**
 * Factor of safety of a circular slip in undrained clay (φ = 0) using sector geometry. Inputs: pline ground surface, center, radius, cu, unit weight, detail (any value for details).
 * @customfunction SLOPE_CIRCLE_COHESIVE
 * @helpurl https://jpbardet.github.io/geoXL/index.html#SLOPE_CIRCLE_COHESIVE
 * @param {any[][]} pline Ground surface polyline (two columns x, y)
 * @param {any[][]} center Circle centre {x, y}
 * @param {any[][]} radius Circle radius (m)
 * @param {any[][]} cu Undrained shear strength (kPa)
 * @param {any[][]} unitweight Unit weight (kN/m³)
 * @param {any[][]} [detail] Optional: any value returns the details table instead of FS only
 * @returns {any[][]} Result (single value, row or table).
 */
function slope_circle_cohesive(pline, center, radius, cu, unitweight, detail) {
  return safe(() => G.slope_circle_cohesive(pts2(pline), point(center), num(radius), num(cu), num(unitweight), opt(detail)));
}

/**
 * Geometry of the soil mass above a circular slip: centroid, entry/exit points and angles, area. Inputs: pl ground polyline, center, radius.
 * @customfunction SLOPE_CIRCLE_POLYLINE
 * @helpurl https://jpbardet.github.io/geoXL/index.html#SLOPE_CIRCLE_POLYLINE
 * @param {any[][]} pl Ground surface polyline (two columns x, y)
 * @param {any[][]} center Circle centre {x, y}
 * @param {any[][]} radius Circle radius
 * @returns {any[][]} Result (single value, row or table).
 */
function slope_circle_polyline(pl, center, radius) {
  return safe(() => G.slope_circle_polyline(pts2(pl), point(center), num(radius)));
}

/**
 * Factor of safety of a circular slip in undrained clay using polygon geometry. Inputs: pline, center, radius, cu, unit weight, detail (any value for details).
 * @customfunction SLOPE_CIRCLE_POLYLINE_COHESIVE
 * @helpurl https://jpbardet.github.io/geoXL/index.html#SLOPE_CIRCLE_POLYLINE_COHESIVE
 * @param {any[][]} pline Ground surface polyline (two columns x, y)
 * @param {any[][]} center Circle centre {x, y}
 * @param {any[][]} radius Circle radius (m)
 * @param {any[][]} cu Undrained shear strength (kPa)
 * @param {any[][]} unitweight Unit weight (kN/m³)
 * @param {any[][]} [detail] Optional: any value returns the details table
 * @returns {any[][]} Result (single value, row or table).
 */
function slope_circle_polyline_cohesive(pline, center, radius, cu, unitweight, detail) {
  return safe(() => G.slope_circle_polyline_cohesive(pts2(pline), point(center), num(radius), num(cu), num(unitweight), opt(detail)));
}

/**
 * Friction circle method (c–φ soil): factor of safety and force polygon (labelled table). Inputs: pline, pt circle centre, rValue radius, coh, phi, g, opt_ (any value: alternative Taylor factor).
 * @customfunction SLOPE_FRICTION_CIRCLE
 * @helpurl https://jpbardet.github.io/geoXL/index.html#SLOPE_FRICTION_CIRCLE
 * @param {any[][]} pline Ground surface polyline (two columns x, y)
 * @param {any[][]} pt Circle centre {x, y}
 * @param {any[][]} rValue Circle radius (m)
 * @param {any[][]} coh Cohesion c (kPa)
 * @param {any[][]} phi Friction angle φ (degrees)
 * @param {any[][]} g Unit weight γ (kN/m³)
 * @param {any[][]} [opt_] Optional: any value uses the alternative Taylor factor
 * @returns {any[][]} Result (single value, row or table).
 */
function slope_friction_circle(pline, pt, rValue, coh, phi, g, opt_) {
  return safe(() => G.slope_friction_circle(pts2(pline), point(pt), num(rValue), num(coh), num(phi), num(g), opt(opt_)));
}

/**
 * Helper for slope_friction_circle: cohesion factor of safety for a trial frictional FS (accepts a column of FS).
 * @customfunction SLOPE_FRICTION_CIRCLE_FCT
 * @helpurl https://jpbardet.github.io/geoXL/index.html#SLOPE_FRICTION_CIRCLE_FCT
 * @param {any[][]} FS Trial frictional factor of safety (accepts a column)
 * @param {any[][]} f Friction angle (radians)
 * @param {any[][]} coh Cohesion c
 * @param {any[][]} w Weight of the sliding mass
 * @param {any[][]} kr Taylor factor k
 * @param {any[][]} rValue Circle radius
 * @param {any[][]} OD Distance used for the line of action of the weight
 * @param {any[][]} DM Lever arm of the cohesion resultant
 * @param {any[][]} alpha Direction of the cohesion resultant (radians)
 * @param {any[][]} Lc Chord length
 * @returns {any[][]} Result (single value, row or table).
 */
function slope_friction_circle_fct(FS, f, coh, w, kr, rValue, OD, DM, alpha, Lc) {
  return safeLifted(() => lift((e_FS, e_f, e_coh, e_w, e_kr, e_rValue, e_OD, e_DM, e_alpha, e_Lc) => G.slope_friction_circle_fct(num(e_FS), num(e_f), num(e_coh), num(e_w), num(e_kr), num(e_rValue), num(e_OD), num(e_DM), num(e_alpha), num(e_Lc)), [FS, f, coh, w, kr, rValue, OD, DM, alpha, Lc], []));
}

/**
 * Taylor friction-circle radius factor k for central angle t (radians). opt = 1 selects the alternative formula.
 * @customfunction SLOPE_FRICTION_CIRCLE_TAYLOR
 * @helpurl https://jpbardet.github.io/geoXL/index.html#SLOPE_FRICTION_CIRCLE_TAYLOR
 * @param {any[][]} t Central angle of the arc (radians)
 * @param {any[][]} opt 1 selects the alternative formula
 * @returns {any[][]} Result (single value, row or table).
 */
function slope_friction_circle_Taylor(t, opt) {
  return safeLifted(() => lift((e_t, e_opt) => G.slope_friction_circle_Taylor(num(e_t), num(e_opt)), [t, opt], []));
}

/**
 * Residual FS − FSc for the friction circle method (same inputs as slope_friction_circle_fct).
 * @customfunction SLOPE_FRICTION_FS
 * @helpurl https://jpbardet.github.io/geoXL/index.html#SLOPE_FRICTION_FS
 * @param {any[][]} FS Trial factor of safety
 * @param {any[][]} f Friction angle (radians)
 * @param {any[][]} coh Cohesion c
 * @param {any[][]} w Weight of the sliding mass
 * @param {any[][]} kr Taylor factor k
 * @param {any[][]} rValue Circle radius
 * @param {any[][]} a Distance used for the line of action of the weight
 * @param {any[][]} DM Lever arm of the cohesion resultant
 * @param {any[][]} alpha Direction of the cohesion resultant (radians)
 * @param {any[][]} ab Chord length
 * @returns {any[][]} Result (single value, row or table).
 */
function slope_friction_FS(FS, f, coh, w, kr, rValue, a, DM, alpha, ab) {
  return safeLifted(() => lift((e_FS, e_f, e_coh, e_w, e_kr, e_rValue, e_a, e_DM, e_alpha, e_ab) => G.slope_friction_FS(num(e_FS), num(e_f), num(e_coh), num(e_w), num(e_kr), num(e_rValue), num(e_a), num(e_DM), num(e_alpha), num(e_ab)), [FS, f, coh, w, kr, rValue, a, DM, alpha, ab], []));
}

/**
 * Log-spiral slip surface: factor of safety, centroid, weight, entry/exit points, resisting and driving moments. Inputs: pl ground polyline, pt spiral centre, rValue, phi, coh, g.
 * @customfunction SLOPE_LOGSPIRAL
 * @helpurl https://jpbardet.github.io/geoXL/index.html#SLOPE_LOGSPIRAL
 * @param {any[][]} pl Ground surface polyline (two columns x, y)
 * @param {any[][]} pt Spiral centre {x, y}
 * @param {any[][]} rValue Initial radius
 * @param {any[][]} phi Friction angle φ (degrees)
 * @param {any[][]} coh Cohesion c (kPa)
 * @param {any[][]} g Unit weight γ (kN/m³)
 * @returns {any[][]} Result (single value, row or table).
 */
function slope_logSpiral(pl, pt, rValue, phi, coh, g) {
  return safe(() => G.slope_logSpiral(pts2(pl), point(pt), num(rValue), num(phi), num(coh), num(g)));
}

/**
 * Alias of SPT_correction_depth: overburden correction CN, capped at 2.
 * @customfunction SPT_DEPTHCORRECTION
 * @helpurl https://jpbardet.github.io/geoXL/index.html#SPT_DEPTHCORRECTION
 * @param {any[][]} s Effective vertical stress σ'v (kPa)
 * @returns {any[][]} Result (single value, row or table).
 */
function SPT_depthCorrection(s) {
  return safeLifted(() => lift((e_s) => G.SPT_correction_depth(num(e_s)), [s], []));
}

/**
 * Friction angle (deg) from SPT N. Optional spa_ = σ'v/pa (default 1), author_: peck (default), schmertman, hatanaka.
 * @customfunction SPT_FRICTIONANGLE
 * @helpurl https://jpbardet.github.io/geoXL/index.html#SPT_FRICTIONANGLE
 * @param {any[][]} N SPT blow count N
 * @param {any[][]} [spa_] Ratio σ'v/pa, optional, default 1
 * @param {any[][]} [author_] "peck" (default), "schmertman" or "hatanaka"
 * @returns {any[][]} Result (single value, row or table).
 */
function SPT_frictionAngle(N, spa_, author_) {
  return safeLifted(() => lift((e_N, e_spa_, e_author_) => G.SPT_frictionAngle(num(e_N), e_spa_, e_author_), [N, spa_, author_], []));
}

/**
 * Staircase table: each x repeated twice with y shifted one row, for step plots. Inputs: x, y columns.
 * @customfunction STEP_TABLE
 * @helpurl https://jpbardet.github.io/geoXL/index.html#STEP_TABLE
 * @param {any[][]} x Column of x values
 * @param {any[][]} y Column of y values
 * @returns {any[][]} Result (single value, row or table).
 */
function Step_table(x, y) {
  return safe(() => G.Step_table(colNum(x), colNum(y)));
}

/**
 * Intersections {x, y} of a closed polygon with the line through M at angle t (radians). Returns {0, 0} if none.
 * @customfunction INTERSECTION_POLYGON_LINE_SLOPE
 * @helpurl https://jpbardet.github.io/geoXL/index.html#INTERSECTION_POLYGON_LINE_SLOPE
 * @param {any[][]} polyline Closed polygon (two columns x, y)
 * @param {any[][]} M Point {x, y} on the line
 * @param {any[][]} t Line angle (radians)
 * @returns {any[][]} Result (single value, row or table).
 */
function intersection_polygon_line_slope(polyline, M, t) {
  return safe(() => G.intersection_polygon_line_slope(pts2(polyline), point(M), num(t)));
}

/**
 * Grain-size analysis: {D10; D30; D60; Cu; Cc; % clay; % silt; % sand; % gravel}. Inputs: d diameters (mm), p percent passing (any order).
 * @customfunction GRAINSIZEANALYSIS
 * @helpurl https://jpbardet.github.io/geoXL/index.html#GRAINSIZEANALYSIS
 * @param {any[][]} d Column of particle diameters (mm)
 * @param {any[][]} p Column of percent passing (%), any order
 * @returns {any[][]} Result (single value, row or table).
 */
function grainsizeanalysis(d, p) {
  return safe(() => G.grainsizeanalysis(colNum(d), colNum(p)));
}

/**
 * Pile capacity in clay by the α method with tip 9·Su·A: {ultimate; shaft; tip}. Inputs: zSu (depth, Su), pa, a tip area, p perimeter, L, author (terzaghi, reese), optional D_ top depth.
 * @customfunction PILE_CAPACITY_ALPHA
 * @helpurl https://jpbardet.github.io/geoXL/index.html#PILE_CAPACITY_ALPHA
 * @param {any[][]} zSu Two-column table {depth (m), Su (kPa)}
 * @param {any[][]} pa Atmospheric pressure (kPa)
 * @param {any[][]} a Pile tip area (m²)
 * @param {any[][]} p Pile perimeter (m)
 * @param {any[][]} L Pile length (m)
 * @param {any[][]} author "terzaghi" or "reese"
 * @param {any[][]} [D_] Depth of the pile top (m), optional, default 0
 * @returns {any[][]} Result (single value, row or table).
 */
function pile_capacity_alpha(zSu, pa, a, p, L, author, D_) {
  return safe(() => G.pile_capacity_alpha(mat(zSu), num(pa), num(a), num(p), num(L), str(author), opt(D_)));
}

/**
 * LCPC (Bustamante–Gianeselli) CPT pile capacity: ultimate, shaft, tip, average qc at tip, average qs, kc. Inputs: table (z, qc, soil), a tip area, p perimeter, L, pile type.
 * @customfunction PILE_CAPACITY_CPT_LCPC
 * @helpurl https://jpbardet.github.io/geoXL/index.html#PILE_CAPACITY_CPT_LCPC
 * @param {any[][]} z_qc_soil Table {depth (m), qc (MPa), soil}
 * @param {any[][]} a Pile tip area (m²)
 * @param {any[][]} p Pile perimeter (m)
 * @param {any[][]} L Pile length (m)
 * @param {any[][]} pile Pile type ("drilled", "cast", "steel", "concrete")
 * @returns {any[][]} Result (single value, row or table).
 */
function pile_capacity_CPT_LCPC(z_qc_soil, a, p, L, pile) {
  return safe(() => G.pile_capacity_CPT_LCPC(mat(z_qc_soil), num(a), num(p), num(L), str(pile)));
}

/**
 * Fellenius β-method pile capacity: {ultimate; shaft; tip}. Inputs: table (z, phi, σ'v, soil), a tip area, p perimeter, L, optional D_ top depth.
 * @customfunction PILE_CAPACITY_FELLENIUS
 * @helpurl https://jpbardet.github.io/geoXL/index.html#PILE_CAPACITY_FELLENIUS
 * @param {any[][]} z_phi_sv_soil Table {depth (m), φ (degrees), σ'v (kPa), soil ("clay", "silt", "sand", "gravel")}
 * @param {any[][]} a Pile tip area (m²)
 * @param {any[][]} p Pile perimeter (m)
 * @param {any[][]} L Pile length (m)
 * @param {any[][]} [D_] Depth of the pile top (m), optional, default 0
 * @returns {any[][]} Result (single value, row or table).
 */
function pile_capacity_Fellenius(z_phi_sv_soil, a, p, L, D_) {
  return safe(() => G.pile_capacity_Fellenius(mat(z_phi_sv_soil), num(a), num(p), num(L), opt(D_)));
}

/**
 * LCPC pressuremeter pile capacity: {ultimate; shaft; tip}. Inputs: table (z, pl, σv, soil), a tip area, p perimeter, L, pile type.
 * @customfunction PILE_CAPACITY_PMT_LCPC
 * @helpurl https://jpbardet.github.io/geoXL/index.html#PILE_CAPACITY_PMT_LCPC
 * @param {any[][]} z_pl_sv_soil Table {depth (m), pl (MPa), σv, soil}
 * @param {any[][]} a Pile tip area (m²)
 * @param {any[][]} p Pile perimeter (m)
 * @param {any[][]} L Pile length (m)
 * @param {any[][]} pile Pile installation type
 * @returns {any[][]} Result (single value, row or table).
 */
function pile_capacity_PMT_LCPC(z_pl_sv_soil, a, p, L, pile) {
  return safe(() => G.pile_capacity_PMT_LCPC(mat(z_pl_sv_soil), num(a), num(p), num(L), str(pile)));
}

/**
 * Meyerhof SPT pile capacity: ultimate, shaft, tip, average N along shaft and at tip. Inputs: zN (depth, N), pa, a tip area, p perimeter, L, disp (any value = displacement pile).
 * @customfunction PILE_CAPACITY_SPT_MEYERHOF
 * @helpurl https://jpbardet.github.io/geoXL/index.html#PILE_CAPACITY_SPT_MEYERHOF
 * @param {any[][]} zN Two-column table {depth (m), N}
 * @param {any[][]} pa Atmospheric pressure (kPa)
 * @param {any[][]} a Pile tip area (m²)
 * @param {any[][]} p Pile perimeter (m)
 * @param {any[][]} L Pile length (m)
 * @param {any[][]} [disp] Optional: any value = displacement (driven) pile
 * @returns {any[][]} Result (single value, row or table).
 */
function pile_capacity_SPT_Meyerhof(zN, pa, a, p, L, disp) {
  return safe(() => G.pile_capacity_SPT_Meyerhof(mat(zN), num(pa), num(a), num(p), num(L), opt(disp)));
}

/**
 * Shaft capacity by the λ method: capacity, λ, average Su and σ'v. Inputs: table (z, Su, σ'v), p perimeter, L (m, or ft if ft is supplied).
 * @customfunction PILE_SHAFT_LAMBDA
 * @helpurl https://jpbardet.github.io/geoXL/index.html#PILE_SHAFT_LAMBDA
 * @param {any[][]} z_Su_sv Table {depth, Su, σ'v}
 * @param {any[][]} p Pile perimeter
 * @param {any[][]} L Pile length (m, or ft if ft is supplied)
 * @param {any[][]} [ft] Optional: any value means lengths are in feet
 * @returns {any[][]} Result (single value, row or table).
 */
function pile_shaft_lambda(z_Su_sv, p, L, ft) {
  return safe(() => G.pile_shaft_lambda(mat(z_Su_sv), num(p), num(L), opt(ft)));
}

/**
 * LCPC unit shaft friction (MPa) from limit pressure pl (MPa) on curve i (1–7). Accepts columns.
 * @customfunction PMT_SHAFTFRICTION
 * @helpurl https://jpbardet.github.io/geoXL/index.html#PMT_SHAFTFRICTION
 * @param {any[][]} pl Limit pressure pl (MPa); accepts a column
 * @param {any[][]} i LCPC curve number 1–7
 * @returns {any[][]} Result (single value, row or table).
 */
function PMT_shaftFriction(pl, i) {
  return safeLifted(() => lift((e_pl, e_i) => G.PMT_shaftFriction(num(e_pl), num(e_i)), [pl, i], []));
}

/**
 * LCPC unit shaft friction (MPa) from limit pressure pl (MPa) on curve "q1" … "q7".
 * @customfunction PMT_PILE_SHAFTFRICTION
 * @helpurl https://jpbardet.github.io/geoXL/index.html#PMT_PILE_SHAFTFRICTION
 * @param {any[][]} pl Limit pressure pl (MPa)
 * @param {any[][]} classi Curve "q1" … "q7"
 * @returns {any[][]} Result (single value, row or table).
 */
function PMT_pile_shaftFriction(pl, classi) {
  return safeLifted(() => lift((e_pl, e_classi) => G.PMT_pile_shaftFriction(num(e_pl), e_classi), [pl, classi], []));
}

/**
 * Ménard shape coefficients λc, λd for L/B. opt_: "v" = λc, "d" = λd, omitted = {λc; λd}. Returns 0 if L/B < 1.
 * @customfunction PMT_SETTLEMENT_LAMBDA
 * @helpurl https://jpbardet.github.io/geoXL/index.html#PMT_SETTLEMENT_LAMBDA
 * @param {any[][]} LB Ratio L/B
 * @param {any[][]} [opt_] Optional: "v" = λc only, "d" = λd only; omit for both
 * @returns {any[][]} Result (single value, row or table).
 */
function PMT_settlement_lambda(LB, opt_) {
  return safe(() => G.PMT_settlement_lambda(num(LB), opt(opt_)));
}

/**
 * Alias of PMT_settlement_lambda: Ménard shape coefficients {λc; λd} for L/B.
 * @customfunction PMT_LAMBDA
 * @helpurl https://jpbardet.github.io/geoXL/index.html#PMT_LAMBDA
 * @param {any[][]} LB Ratio L/B
 * @returns {any[][]} Result (single value, row or table).
 */
function PMT_lambda(LB) {
  return safe(() => G.PMT_settlement_lambda(num(LB), "a"));
}

/**
 * Ménard equivalent moduli Ec, Ed and limit pressure from 16 sub-layers below a footing (labelled table + layer values). Inputs: yEpl (depth, E, pl), B, D.
 * @customfunction PMT_MODULI
 * @helpurl https://jpbardet.github.io/geoXL/index.html#PMT_MODULI
 * @param {any[][]} yEpl Table {depth (m), E, pl}
 * @param {any[][]} B Footing width (m)
 * @param {any[][]} D Footing depth (m)
 * @returns {any[][]} Result (single value, row or table).
 */
function PMT_moduli(yEpl, B, D) {
  return safe(() => G.PMT_moduli(mat(yEpl), num(B), num(D)));
}

/**
 * Eurocode 7 pressuremeter bearing pressure: {qu; ple; k}. Inputs: z, pl columns, B, D, L, soil ("clay silt", "sand gravel", "chalk"), g unit weight.
 * @customfunction PMT_BEARING_EUROCODE
 * @helpurl https://jpbardet.github.io/geoXL/index.html#PMT_BEARING_EUROCODE
 * @param {any[][]} z Column of test depths (m)
 * @param {any[][]} pl Column of limit pressures pl (MPa)
 * @param {any[][]} B Footing width (m)
 * @param {any[][]} D Footing depth (m)
 * @param {any[][]} L Footing length (m)
 * @param {any[][]} soil "clay silt", "sand gravel" or "chalk"
 * @param {any[][]} g Unit weight above footing level
 * @returns {any[][]} Result (single value, row or table).
 */
function PMT_bearing_Eurocode(z, pl, B, D, L, soil, g) {
  return safe(() => G.PMT_bearing_Eurocode(colNum(z), colNum(pl), num(B), num(D), num(L), str(soil), num(g)));
}

/**
 * Eurocode 7 pressuremeter bearing factor k. Inputs: B, L, D embedment, pl limit pressure (MPa), soil_ ("clay silt", "sand gravel", "chalk").
 * @customfunction PMT_BEARING_EUROCODE_K
 * @helpurl https://jpbardet.github.io/geoXL/index.html#PMT_BEARING_EUROCODE_K
 * @param {any[][]} B Footing width (m)
 * @param {any[][]} L Footing length (m)
 * @param {any[][]} D Embedment (m)
 * @param {any[][]} pl Limit pressure pl (MPa)
 * @param {any[][]} soil_ "clay silt", "sand gravel" or "chalk"
 * @returns {any[][]} Result (single value, row or table).
 */
function PMT_bearing_eurocode_k(B, L, D, pl, soil_) {
  return safeLifted(() => lift((e_B, e_L, e_D, e_pl, e_soil_) => G.PMT_bearing_eurocode_k(num(e_B), num(e_L), num(e_D), num(e_pl), e_soil_), [B, L, D, pl, soil_], []));
}

/**
 * Ménard settlement {sc; sd; total} from given moduli. Inputs: q net pressure, Ec, Ed, B, LB = L/B, ac, ad rheological factors, optional B0_ reference width (default 1).
 * @customfunction PMT_SETTLEMENT_BASIC
 * @helpurl https://jpbardet.github.io/geoXL/index.html#PMT_SETTLEMENT_BASIC
 * @param {any[][]} q Net applied pressure
 * @param {any[][]} Ec Volumetric (spherical) modulus Ec, same unit as q
 * @param {any[][]} Ed Deviatoric modulus Ed
 * @param {any[][]} B Footing width (m)
 * @param {any[][]} LB Ratio L/B
 * @param {any[][]} ac Rheological factor α for the volumetric term
 * @param {any[][]} ad Rheological factor α for the deviatoric term
 * @param {any[][]} [B0_] Reference width (m), optional, default 1
 * @returns {any[][]} Result (single value, row or table).
 */
function PMT_settlement_basic(q, Ec, Ed, B, LB, ac, ad, B0_) {
  return safe(() => G.PMT_settlement_basic(num(q), num(Ec), num(Ed), num(B), num(LB), num(ac), num(ad), opt(B0_)));
}

/**
 * Ménard settlement from a pressuremeter profile (labelled table: total, volumetric, deviatoric, α, λ, moduli). Inputs: q, yEpl (depth, E, pl), B, D, L, soil, ft_ (units flag).
 * @customfunction PMT_SETTLEMENT
 * @helpurl https://jpbardet.github.io/geoXL/index.html#PMT_SETTLEMENT
 * @param {any[][]} q Net applied pressure
 * @param {any[][]} yEpl Table {depth (m), E, pl}
 * @param {any[][]} B Footing width (m)
 * @param {any[][]} D Footing depth (m)
 * @param {any[][]} L Footing length (m)
 * @param {any[][]} soil "peat", "clay", "silt", "sand", "gravel" or "rock"
 * @param {any[][]} [ft_] Optional units flag: omit for B0 = 0.6 m; any value for B0 = 2 ft
 * @returns {any[][]} Result (single value, row or table).
 */
function PMT_settlement(q, yEpl, B, D, L, soil, ft_) {
  return safe(() => G.PMT_settlement(num(q), mat(yEpl), num(B), num(D), num(L), str(soil), opt(ft_)));
}

/**
 * Shear and moment from a piecewise-linear pressure profile. Columns: depth, pressure, shear, moment. k = subdivisions per segment.
 * @customfunction SHEARMOMENT
 * @helpurl https://jpbardet.github.io/geoXL/index.html#SHEARMOMENT
 * @param {any[][]} d Column of depths (m)
 * @param {any[][]} p Column of pressures (kPa)
 * @param {any[][]} k Subdivisions per segment
 * @returns {any[][]} Result (single value, row or table).
 */
function shearMoment(d, p, k) {
  return safe(() => G.shearMoment(colNum(d), colNum(p), num(k)));
}

/**
 * Shear and moment diagram without header (depth, pressure, shear, moment). Inputs: d, p columns, optional n subdivisions (default 20).
 * @customfunction BEAM
 * @helpurl https://jpbardet.github.io/geoXL/index.html#BEAM
 * @param {any[][]} d Column of depths (m)
 * @param {any[][]} p Column of pressures (kPa)
 * @param {any[][]} [n] Subdivisions per segment, optional, default 20
 * @returns {any[][]} Result (single value, row or table).
 */
function beam(d, p, n) {
  return safe(() => G.shearMoment(colNum(d), colNum(p), opt(n) === undefined ? 20 : num(opt(n))));
}

/**
 * Shear and moment diagram with header (depth, pressure, shear, moment). Inputs: d, p columns, optional n subdivisions (default 20).
 * @customfunction PRESSURE_SHEAR_MOMENT
 * @helpurl https://jpbardet.github.io/geoXL/index.html#PRESSURE_SHEAR_MOMENT
 * @param {any[][]} d Column of depths (m)
 * @param {any[][]} p Column of pressures (kPa)
 * @param {any[][]} [n] Subdivisions per segment, optional, default 20
 * @returns {any[][]} Result (single value, row or table).
 */
function pressure_shear_moment(d, p, n) {
  return safe(() => [SM_HDR, ...G.shearMoment(colNum(d), colNum(p), opt(n) === undefined ? 20 : num(opt(n)))]);
}

/**
 * Shear and moment diagram with an anchor force Ta at depth da (sheet-pile walls). Inputs: d, p columns, da, Ta, optional n subdivisions (default 20).
 * @customfunction PRESSURE_SHEAR_MOMENT_ANCHOR
 * @helpurl https://jpbardet.github.io/geoXL/index.html#PRESSURE_SHEAR_MOMENT_ANCHOR
 * @param {any[][]} d Column of depths (m)
 * @param {any[][]} p Column of pressures (kPa)
 * @param {any[][]} da Depth of the anchor (m)
 * @param {any[][]} Ta Anchor force (kN/m)
 * @param {any[][]} [n] Subdivisions per segment, optional, default 20
 * @returns {any[][]} Result (single value, row or table).
 */
function pressure_shear_moment_anchor(d, p, da, Ta, n) {
  return safe(() => G.pressure_shear_moment_anchor(colNum(d), colNum(p), num(da), num(Ta), opt(n)));
}

/**
 * Allowable bearing pressure from a single SPT N value. Inputs: N, B, D (m), author: bowles, meyerhof, terzaghi.
 * @customfunction SPT_BEARING_DIRECT
 * @helpurl https://jpbardet.github.io/geoXL/index.html#SPT_BEARING_DIRECT
 * @param {any[][]} N SPT blow count N (average below the footing)
 * @param {any[][]} B Footing width (m)
 * @param {any[][]} D Footing depth (m)
 * @param {any[][]} author "bowles", "meyerhof" or "terzaghi"
 * @returns {any[][]} Result (single value, row or table).
 */
function SPT_bearing_direct(N, B, D, author) {
  return safeLifted(() => lift((e_N, e_B, e_D, e_author) => G.SPT_bearing_direct(num(e_N), num(e_B), num(e_D), e_author), [N, B, D, author], []));
}

/**
 * ASTM D2487 group symbol and group name {symbol, name}. Inputs: % passing No. 4 and No. 200, optional Cu, Cc, LL, PL, LL oven-dried/LL not dried (blank = inorganic).
 * @customfunction CLASSIFY_ASTM2487
 * @helpurl https://jpbardet.github.io/geoXL/index.html#CLASSIFY_ASTM2487
 * @param {any[][]} pf4 Percent passing No. 4 sieve (%)
 * @param {any[][]} pf200 Percent passing No. 200 sieve (%)
 * @param {any[][]} [cu_] Coefficient of uniformity Cu = D60/D10, optional
 * @param {any[][]} [cc_] Coefficient of curvature Cc = D30²/(D10·D60), optional
 * @param {any[][]} [ll_] Liquid limit LL (%), optional
 * @param {any[][]} [pl_] Plastic limit PL (%), optional
 * @param {any[][]} [ratio_] LL oven-dried / LL not dried, optional (below 0.75 = organic; blank = inorganic)
 * @returns {any[][]} Result (single value, row or table).
 */
function classify_ASTM2487(pf4, pf200, cu_, cc_, ll_, pl_, ratio_) {
  return safeLifted(() => liftRows((e_pf4, e_pf200, e_cu_, e_cc_, e_ll_, e_pl_, e_ratio_) => G.classify_ASTM2487(num(e_pf4), num(e_pf200), e_cu_, e_cc_, e_ll_, e_pl_, e_ratio_), [pf4, pf200, cu_, cc_, ll_, pl_, ratio_], []));
}

// ---------------------------------------------------------------- registration
const GEOXL_FUNCTIONS = {VERSION: version, BEARING_CAPACITY: bearing_capacity, BEARING_DEPTHFACTORS: bearing_depthFactors, BEARING_FACTORS: bearing_factors, BEARING_INCLINATIONFACTORS: bearing_inclinationFactors, BEARING_NQ: bearing_Nq, BEARING_SHAPEFACTORS: bearing_shapeFactors, BEARING_SKEMPTON: bearing_Skempton, CLASSIFY_AASHTO: classify_AASHTO, ASTM_TABLE: astm_table, CLASSIFY_ASTM: classify_ASTM, CLASSIFY_PLASTICITY: classify_Plasticity, CLASSIFY_USC: classify_USC, CONSOLIDATION_TEST: consolidation_test, CONSOLIDATION_TV: consolidation_Tv, CONSOLIDATION_U: consolidation_U, CONSOLIDATION_UX: consolidation_Ux, CONSOLIDATION_U_APPROXIMATE: consolidation_U_approximate, CONSOLIDATION_U_VARIABLE: consolidation_U_variable, COSD: cosd, COULOMB_COHESIVEFORCES: Coulomb_cohesiveforces, COULOMB_FORCES: Coulomb_forces, COULOMB_K: coulomb_K, COULOMB_MAXIMUM: Coulomb_maximum, COULOMB_THETA: coulomb_theta, CUTLINE: cutline, ELASTIC_CIRCULAR_RADIUS: elastic_circular_radius, ELASTIC_CIRCULAR_STRESS: elastic_circular_stress, ELASTIC_EMBANKMENT: elastic_embankment, ELASTIC_JANBU_MU0: elastic_Janbu_mu0, ELASTIC_JANBU_MU1: elastic_Janbu_mu1, ELASTIC_LINELOAD: elastic_lineload, ELASTIC_LINELOAD_HORIZONTAL: elastic_lineload_horizontal, ELASTIC_NEWMARK_FACTOR: elastic_Newmark_factor, ELASTIC_NEWMARK_STRESS: elastic_Newmark_stress, ELASTIC_POINTLOAD: elastic_pointload, ELASTIC_SECTOR: elastic_sector, ELASTIC_SECTOR_ROW: elastic_sector_row, ELASTIC_SECTOR_SLICE: elastic_sector_slice, ELASTIC_STRIPLOAD: elastic_stripload, ELASTIC_TRIANGLE: elastic_triangle, INTERPOLATE: Interpolate, INTERSECTION_POLYGON_LINE: intersection_polygon_line, INTERSECTION_POLYGON_RAY: intersection_polygon_ray, MINIMUM: minimum, POLYGON_CENTROIDAREA: polygon_centroidArea, POLYGON_WEIGHTMOMENT: polygon_weightMoment, PRESSURE_FORCEMOMENT: pressure_forceMoment, RANKINE_K: Rankine_K, RANKINE_STRESS: Rankine_stress, SIND: sind, SPT_BEARING: SPT_bearing, SPT_BURLANDBURBIDGE_SETTLEMENT: SPT_BurlandBurbidge_settlement, SPT_CORRECTION_DEPTH: SPT_correction_depth, SPT_CORRELATION_FRICTION: SPT_Correlation_friction, STRESS_FORCEMOMENT: stress_ForceMoment, STRESS_MOHR: Stress_Mohr, STRESS_MOHR_FAILURE: stress_Mohr_failure, STRESS_MOHR_POLE: stress_Mohr_pole, STRESS_MOHRCIRCLE: stress_Mohrcircle, TAND: tand, VECTOR_LINE: vector_line, WATER_PRESSURE: water_pressure, ELASTIC_STRIPLOAD_XY: elastic_stripload_xy, CIRCLE: circle, CLASSIFY_CPT_LCPC: classify_CPT_LCPC, CLASSIFY_PMT_LCPC: classify_PMT_LCPC, CPT_AVERAGE_QC: CPT_average_qc, CPT_BEARING_MEYERHOF_K: CPT_bearing_Meyerhof_k, CPT_BEARING_ROBERTSON_K: CPT_bearing_Robertson_k, CPT_FRICTION_ANGLE: CPT_friction_angle, CPT_OLSEN_LAYERS: CPT_Olsen_layers, CPT_OLSEN_SCN: CPT_Olsen_SCN, CPT_OLSEN_SCN_ROW: CPT_Olsen_SCN_row, CPT_ROBERTSON_CLASSIFICATION_CHART: CPT_Robertson_classification_chart, CPT_ROBERTSON_LAYERS: CPT_Robertson_layers, CPT_ROBERTSON_SBT: CPT_Robertson_SBT, CPT_ROBERTSON_SBT_ROW: CPT_Robertson_SBT_row, CPT_SETTLEMENT: CPT_settlement, CPT_SETTLEMENT_IZ: CPT_settlement_Iz, DISTANCE_POINT_POLYLINE: distance_point_polyline, INTERPOLATE_TABLE: interpolate_table, INTERSECTION_POLYLINE_CIRCLE: intersection_polyline_circle, INTERSECTION_POLYLINE_POLYLINE: intersection_polyline_polyline, INTERSECTION_SEGMENT_POLYLINE: intersection_segment_polyline, INTERSECTION_SPIRAL_SEGMENT: intersection_spiral_segment, LATERALPILE_BROMS_COHESIONLESS: lateralpile_Broms_cohesionless, LATERALPILE_BROMS_COHESIVE: lateralpile_Broms_cohesive, LATERALPILE_COHESIONLESS_REACTION_SHEAR_MOMENT: lateralpile_cohesionless_reaction_shear_moment, LATERALPILE_COHESIVE_REACTION_SHEAR_MOMENT: lateralpile_cohesive_reaction_shear_moment, LOGSPIRAL: logSpiral, PILE_CAPACITY_CPT_SCHMERTMAN: pile_capacity_CPT_Schmertman, PILE_SHAFT_ALPHA: pile_shaft_alpha, PILE_SIDEFRICTION_ALPHA: pile_sidefriction_alpha, PILE_SIDEFRICTION_LAMBDA: pile_sidefriction_lambda, PILE_TIP_MEYERHOF: pile_tip_Meyerhof, PILE_TIP_MEYERHOF_FACTOR: pile_tip_meyerhof_factor, PILE_TIP_VESIC: pile_tip_vesic, PILE_TIP_VESIC_FACTOR: pile_tip_vesic_factor, PMT_ALPHA: PMT_alpha, PMT_BEARING_LIMITPRESSURE: PMT_bearing_limitPressure, PMT_BEARING_MENARD: PMT_bearing_Menard, PMT_BEARING_MENARD_K: PMT_bearing_Menard_k, PMT_LAYERS: PMT_layers, STEPPROFILE: stepProfile, SPT_STEPPROFILE: SPT_stepProfile, CPT_STEPPROFILE: CPT_stepProfile, PMT_STAIRCASE: PMT_staircase, POLYGON_AREA_CENTROID: polygon_area_centroid, POLYGON_INSIDE: Polygon_inside, POLYLINE_SUBDIVIDE: polyline_subdivide, SETTLEMENT_SCHMERTMAN_FACTOR: settlement_Schmertman_factor, SLOPE_CIRCLE_COHESIVE: slope_circle_cohesive, SLOPE_CIRCLE_POLYLINE: slope_circle_polyline, SLOPE_CIRCLE_POLYLINE_COHESIVE: slope_circle_polyline_cohesive, SLOPE_FRICTION_CIRCLE: slope_friction_circle, SLOPE_FRICTION_CIRCLE_FCT: slope_friction_circle_fct, SLOPE_FRICTION_CIRCLE_TAYLOR: slope_friction_circle_Taylor, SLOPE_FRICTION_FS: slope_friction_FS, SLOPE_LOGSPIRAL: slope_logSpiral, SPT_DEPTHCORRECTION: SPT_depthCorrection, SPT_FRICTIONANGLE: SPT_frictionAngle, STEP_TABLE: Step_table, INTERSECTION_POLYGON_LINE_SLOPE: intersection_polygon_line_slope, GRAINSIZEANALYSIS: grainsizeanalysis, PILE_CAPACITY_ALPHA: pile_capacity_alpha, PILE_CAPACITY_CPT_LCPC: pile_capacity_CPT_LCPC, PILE_CAPACITY_FELLENIUS: pile_capacity_Fellenius, PILE_CAPACITY_PMT_LCPC: pile_capacity_PMT_LCPC, PILE_CAPACITY_SPT_MEYERHOF: pile_capacity_SPT_Meyerhof, PILE_SHAFT_LAMBDA: pile_shaft_lambda, PMT_SHAFTFRICTION: PMT_shaftFriction, PMT_PILE_SHAFTFRICTION: PMT_pile_shaftFriction, PMT_SETTLEMENT_LAMBDA: PMT_settlement_lambda, PMT_LAMBDA: PMT_lambda, PMT_MODULI: PMT_moduli, PMT_BEARING_EUROCODE: PMT_bearing_Eurocode, PMT_BEARING_EUROCODE_K: PMT_bearing_eurocode_k, PMT_SETTLEMENT_BASIC: PMT_settlement_basic, PMT_SETTLEMENT: PMT_settlement, SHEARMOMENT: shearMoment, BEAM: beam, PRESSURE_SHEAR_MOMENT: pressure_shear_moment, PRESSURE_SHEAR_MOMENT_ANCHOR: pressure_shear_moment_anchor, SPT_BEARING_DIRECT: SPT_bearing_direct, CLASSIFY_ASTM2487: classify_ASTM2487};
if (typeof CustomFunctions !== "undefined") { for (const [id, f] of Object.entries(GEOXL_FUNCTIONS)) CustomFunctions.associate(id, f); }
if (typeof module !== "undefined") module.exports = { G, GEOXL_FUNCTIONS };
