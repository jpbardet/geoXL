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
G.classify_ASTM = (list, pf4, pf200, Cu_, Cc_, LL_, PL_, org_) => {
  const o = (v) => (v === undefined ? 0 : num(v));
  const organic = o(org_), Cu = o(Cu_), Cc = o(Cc_), LL = o(LL_), PL = o(PL_);
  const sand = pf4 - pf200, gravel = 100 - pf4;
  const U = G.classify_USC(pf4, pf200, Cu, Cc, LL, PL, organic);
  const USC = organic > 0 && (U === "ML" || U === "MH") ? "O" + U : organic > 0 && U === "MC" ? "O" + U : U;
  const i = pf200 > 85 ? 1 : pf200 > 70 ? (sand >= gravel ? 2 : 3) : sand >= gravel ? (gravel > 15 ? 4 : 5) : (sand > 15 ? 6 : 7);
  const j = pf200 > 50 ? i : sand < 15 ? 1 : 2;
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
