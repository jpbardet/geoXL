// Node test suite: textbook checks + a smoke run of every function, called the way Excel calls them.
const { GEOXL_FUNCTIONS: F } = require("../docs/functions.js");
const M = (v) => (Array.isArray(v) ? (Array.isArray(v[0]) ? v : v.map((x) => [x])) : [[v]]);   // value/column -> Excel matrix
const S = (x) => [[x]];
const isErr = (c) => c && typeof c === "object" && "xlError" in c;
let pass = 0, fail = 0;
function near(label, got, want, tol = 1e-3) {
  const ok = typeof got === "number" && Math.abs(got - want) <= tol * Math.max(1, Math.abs(want));
  ok ? pass++ : fail++;
  if (!ok) console.log("FAIL", label, "got", got, "want", want);
}
function eq(label, got, want) { const ok = got === want; ok ? pass++ : fail++; if (!ok) console.log("FAIL", label, "got", got, "want", want); }

// ---------------------------------------------------------------- textbook values
let r = F.BEARING_FACTORS(S(30), S("prandtl"));
near("Prandtl Nc(30)", r[0][0], 30.14, 2e-3); near("Prandtl Nq(30)", r[1][0], 18.40, 2e-3);
near("Hansen Ng(30)", F.BEARING_FACTORS(S(30), S("hansen"))[2][0], 15.07, 2e-3);
near("Vesic Ng(30)", F.BEARING_FACTORS(S(30), S("vesic"))[2][0], 22.40, 2e-3);
near("Terzaghi Nq(30)", F.BEARING_FACTORS(S(30), S("terzaghi"))[1][0], 22.46, 2e-3);
near("Coulomb Ka 30/90/0/0", F.COULOMB_K(S(30), S(90), S(0), S(0))[0][0], 1 / 3);
near("Coulomb Kp 30/90/0/0", F.COULOMB_K(S(30), S(90), S(0), S(0), S(1))[0][0], 3);
near("Rankine Ka 30", F.RANKINE_K(S(30), S(0))[0][0], 1 / 3);
r = F.COULOMB_MAXIMUM(S(0), S(0), S(30), S(90), S(0), S(0), S(18), S(5));
near("Coulomb max slip angle", r[0][0], 60, 2e-2); near("Coulomb max K", r[4][0], 1 / 3, 2e-3);
near("Coulomb theta closed form", F.COULOMB_THETA(S(30), S(90), S(0), S(0))[0][0], 60, 1e-3);
near("U(Tv=0.197)", F.CONSOLIDATION_U(S(0.197))[0][0], 0.5, 2e-3);
near("U(Tv=0.848)", F.CONSOLIDATION_U(S(0.848))[0][0], 0.9, 2e-3);
near("Tv(U=0.5)", F.CONSOLIDATION_TV(S(0.5))[0][0], 0.19635, 1e-4);
near("Tv(U=0.9)", F.CONSOLIDATION_TV(S(0.9))[0][0], 0.848, 1e-3);
near("Boussinesq sz under point", F.ELASTIC_POINTLOAD(S(0), S(1))[0][1], 3 / (2 * Math.PI));
near("circle centre sz r=z", F.ELASTIC_CIRCULAR_STRESS(S(1), S(1))[0][1], 1 - 1 / 2 ** 1.5);
near("strip centre sz B=2z", F.ELASTIC_STRIPLOAD(S(2), S(0), S(1))[0][1], (Math.PI / 2 + 1) / Math.PI);
near("Newmark m=n=1", F.ELASTIC_NEWMARK_FACTOR(S(1), S(1))[0][0], 0.1752, 1e-3);
r = F.WATER_PRESSURE(M([1, 3, 5]), S(2));
near("u(1)", r[0][0], 0); near("u(3)", r[1][0], 9.81); near("u(5)", r[2][0], 29.43);
near("Interpolate", F.INTERPOLATE(S(2.5), M([1, 2, 3]), M([10, 20, 30]))[0][0], 25);
near("Interpolate decreasing x", F.INTERPOLATE(S(2.5), M([3, 2, 1]), M([30, 20, 10]))[0][0], 25);
eq("Interpolate outside -> #N/A", F.INTERPOLATE(S(5), M([1, 2, 3]), M([10, 20, 30]))[0][0].xlError, "notAvailable");
near("Interpolate clamp", F.INTERPOLATE(S(5), M([1, 2, 3]), M([10, 20, 30]), null, S("clamp"))[0][0], 30);
near("Interpolate logx", F.INTERPOLATE(S(10), M([1, 100]), M([0, 2]), S("logx"))[0][0], 1);
r = F.INTERPOLATE(M([1.5, 2.5]), M([1, 2, 3]), M([10, 20, 30]));
near("Interpolate column 1", r[0][0], 15); near("Interpolate column 2", r[1][0], 25);
near("CN(100 kPa)", F.SPT_CORRECTION_DEPTH(S(100))[0][0], 1);
r = F.POLYGON_AREA_CENTROID([[0, 0], [2, 0], [2, 2], [0, 2]]);
near("square centroid x", r[0][1], 1); near("square area", r[1][1], 4);
r = F.STRESS_MOHR_POLE(S(100), S(0), S(0), S(50));
near("Mohr centre", r[0][1], 75); near("Mohr radius", r[1][1], 25);
r = F.GRAINSIZEANALYSIS(M([4.75, 2, 0.85, 0.425, 0.25, 0.15, 0.075]), M([95, 80, 60, 40, 25, 15, 8]));
near("D60 (sieve data in decreasing order)", r[2][0], 0.85);
r = F.PRESSURE_SHEAR_MOMENT(M([0, 2]), M([10, 10]));
near("uniform load shear", r[r.length - 1][2], 20); near("uniform load moment", r[r.length - 1][3], 20);
r = F.SHEARMOMENT(M([0, 3]), M([0, 30]), S(20));
near("triangular load shear", r[r.length - 1][2], 45); near("triangular load moment", r[r.length - 1][3], 45);
r = F.CLASSIFY_ASTM2487(S(85), S(8), S(7), S(1.5), S(30), S(24));
eq("ASTM SW-SM", r[0][0], "SW-SM"); eq("ASTM name", r[0][1], "Well-graded sand with silt and gravel");
r = F.CLASSIFY_ASTM2487(M([85, 95]), M([8, 80]), null, null, M([30, 60]), M([24, 30]), M(["", 0.6]));
eq("ASTM column row 2", r[1][1], "Organic clay with sand");
near("Skempton DB=0 square", F.BEARING_SKEMPTON(S(0), S(1))[0][0], 6.2304, 1e-4);
near("cosd(60)", F.COSD(S(60))[0][0], 0.5);
near("Janbu mu0 DB=0.1", F.ELASTIC_JANBU_MU0(S(0.1), S(1))[0][0], 1);

// ---------------------------------------------------------------- smoke run: every function, realistic inputs
const prof = [[0, 20], [2, 25], [4, 30], [6, 35], [8, 40], [10, 45], [12, 50], [14, 55], [16, 60], [18, 65], [20, 70], [22, 75], [24, 80], [26, 85], [28, 90], [30, 95]];
const zqc = prof.map(([z, q]) => [z, q / 4]);                       // depth, qc (MPa)
const cpt = prof.map(([z, q]) => [z + 0.5, q / 4, q / 400]);         // depth, qc, fs (MPa)
const cptSoil = Array.from({ length: 60 }, (_, i) => [0.25 + 0.5 * i, 2 + 0.2 * i, 0.02 + 0.001 * i, "sand"]);   // CPT every 0.5 m
const ground = [[-10, 0], [0, 0], [10, 5], [30, 5]];
const sq = [[0, 0], [2, 0], [2, 2], [0, 2]];
const pyEpl = [[0, 5, 0.5], [2, 8, 0.8], [4, 10, 1.0], [6, 12, 1.2], [8, 14, 1.4], [10, 16, 1.6], [12, 18, 1.8], [14, 20, 2.0], [20, 25, 2.5]];
const T = [0.1, 0.25, 0.5, 1, 2, 4, 8, 15, 30, 60, 120, 240, 480, 1440];
const Xr = [10, 9.95, 9.9, 9.82, 9.7, 9.55, 9.35, 9.15, 8.95, 8.85, 8.8, 8.77, 8.75, 8.72];
const calls = {
  VERSION: [], BEARING_CAPACITY: [S(10), S(30), S(20), S(18), S(2), S(10), S(1), S(0), S(100), S("vesic")],
  BEARING_DEPTHFACTORS: [S(0.5), S(30), S("hansen")], BEARING_FACTORS: [S(30), S("ec7")], BEARING_INCLINATIONFACTORS: [S(10), S(100), S(0.2), S(30), S(10), S("vesic")],
  BEARING_NQ: [S(30), S("vesic")], BEARING_SHAPEFACTORS: [S(0.2), S(30), S("hansen")], BEARING_SKEMPTON: [S(1), S(0.5)],
  CLASSIFY_AASHTO: [S(60), S(45), S(40), S(35), S(20)], ASTM_TABLE: [], CLASSIFY_ASTM: [[["SW", "Well-graded sand"], ["SW", "Well-graded sand with gravel"]], S(85), S(3), S(7), S(2), S(0), S(0)],
  CLASSIFY_PLASTICITY: [S(40), S(20)], CLASSIFY_USC: [S(85), S(8), S(7), S(1.5), S(30), S(24)],
  CONSOLIDATION_TEST: [M(T), M(Xr), S(0.2), S(5), S(20)], CONSOLIDATION_TV: [S(0.7)], CONSOLIDATION_U: [S(0.3), S("triangular_top")],
  CONSOLIDATION_UX: [M([0, 0.25, 0.5, 0.75, 1]), S(0.2)], CONSOLIDATION_U_APPROXIMATE: [S(0.3)], CONSOLIDATION_U_VARIABLE: [S(0.5), S(0.2)],
  COSD: [S(30)], COULOMB_COHESIVEFORCES: [S(1), S(30), S(90), S(0), S(10), S(18), S(5), S(5)], COULOMB_FORCES: [S(1), S(30), S(90), S(0), S(10), S(18), S(5)],
  COULOMB_K: [S(30), S(90), S(10), S(15)], COULOMB_MAXIMUM: [S(0), S(0), S(30), S(90), S(10), S(15), S(18), S(5)], COULOMB_THETA: [S(30), S(90), S(10), S(15)],
  CUTLINE: [M([0, 1, 2, 3]), M([0, 1, 0, 1]), S(0), S(0.5)], ELASTIC_CIRCULAR_RADIUS: [S(0.5)], ELASTIC_CIRCULAR_STRESS: [S(1), S(2)],
  ELASTIC_EMBANKMENT: [S(1), S(2)], ELASTIC_JANBU_MU0: [S(2), S(5)], ELASTIC_JANBU_MU1: [S(5), S("circle")], ELASTIC_LINELOAD: [S(1), S(2)],
  ELASTIC_LINELOAD_HORIZONTAL: [S(1), S(2)], ELASTIC_NEWMARK_FACTOR: [S(0.5), S(2)], ELASTIC_NEWMARK_STRESS: [S(1), S(1), S(2), sq],
  ELASTIC_POINTLOAD: [S(1), S(2)], ELASTIC_SECTOR: [S(1), S(1), S(1), sq], ELASTIC_SECTOR_ROW: [[[0.5, -1, -1, 1, -1, 1, 1, -1, 1]]],
  ELASTIC_SECTOR_SLICE: [S(1), S(2), S(1), S(0.25)], ELASTIC_STRIPLOAD: [S(2), S(1), S(1)], ELASTIC_TRIANGLE: [S(2), S(1), S(1)],
  INTERPOLATE: [S(2.2), M([1, 2, 3]), M([1, 4, 9])], INTERSECTION_POLYGON_LINE: [sq, [[1, 1]], S(0.3)], INTERSECTION_POLYGON_RAY: [sq, [[1, 1]], S(0.3)],
  MINIMUM: [S(0), S(3)], POLYGON_CENTROIDAREA: [sq], POLYGON_WEIGHTMOMENT: [sq, S(20), [[3, 0]]], PRESSURE_FORCEMOMENT: [M([0, 2, 4]), M([0, 20, 40]), S(0)],
  RANKINE_K: [S(30), S(10)], RANKINE_STRESS: [S(30), S(0), S(90), S(10), S(100)], SIND: [S(30)], SPT_BEARING: [M([0, 1, 2, 3, 4]), M([10, 12, 15, 18, 20]), S(2), S(1), S("bowles")],
  SPT_BURLANDBURBIDGE_SETTLEMENT: [M([0, 1, 2, 3, 4]), M([10, 12, 15, 18, 20]), S(150), S(2)], SPT_CORRECTION_DEPTH: [S(50)],
  SPT_CORRELATION_FRICTION: [S(20), S("hatanaka"), S(100)], STRESS_FORCEMOMENT: [[[0, 0]], [[0, 5]], [[10, 2]], [[20, 4]]],
  STRESS_MOHR: [S(100), S(20), S(30), S(80)], STRESS_MOHR_FAILURE: [S(100), S(0), S(0), S(30), S(0)], STRESS_MOHR_POLE: [S(100), S(20), S(30), S(50)],
  STRESS_MOHRCIRCLE: [S(75), S(25)], TAND: [S(45)], VECTOR_LINE: [[[0, 0]], [[0, 5]], [[10, 2]], [[20, 4]], S(4), S(0.1)],
  WATER_PRESSURE: [S(5), S(2)], ELASTIC_STRIPLOAD_XY: [S(2), S(1), S(1)], CIRCLE: [[[0, 0]], S(1)],
  CLASSIFY_CPT_LCPC: [S(10), S("sand"), S("concrete")], CLASSIFY_PMT_LCPC: [S(1.5), S("sand"), S("steel")],
  CPT_AVERAGE_QC: [M(zqc.map((r) => r[0])), M(zqc.map((r) => r[1])), S(2), S(2), S("meyerhof")], CPT_BEARING_MEYERHOF_K: [S(2), S(1), S(5)],
  CPT_BEARING_ROBERTSON_K: [S(2), S(1), S("square"), S("dense")], CPT_FRICTION_ANGLE: [S(5), S(10000), S(18), S("robertson")],
  CPT_OLSEN_LAYERS: [cpt.map((r, i) => [r[0], 0, 0, 0, 0, i < 8 ? 1 : 2])], CPT_OLSEN_SCN: [cpt, S(100), S(0.5), S(18), S(9.81), S(2)],
  CPT_OLSEN_SCN_ROW: [S(50), S(1)], CPT_ROBERTSON_CLASSIFICATION_CHART: [S(1), S("C")],
  CPT_ROBERTSON_LAYERS: [cpt.map((r, i) => [r[0], 0, 0, 0, 0, i < 8 ? 5 : 6, 0, 0, i < 5 ? 5 : 6])],
  CPT_ROBERTSON_SBT: [cpt.map(([z, q, f]) => [z, q * 1000, f * 1000]), S(100), S(0.5), S(18), S(9.81), S(2)], CPT_ROBERTSON_SBT_ROW: [S(50), S(1)],
  CPT_SETTLEMENT: [zqc, S(2.5), S(2), S(1), S(2), S(0.15), S(0.018)], CPT_SETTLEMENT_IZ: [S(2), S(2), S(1), S(2), S(150), S(18)],
  DISTANCE_POINT_POLYLINE: [[[1, 3]], ground], INTERPOLATE_TABLE: [S(1.5), S(1.5), [["", 1, 2], [1, 10, 20], [2, 30, 40]]],
  INTERSECTION_POLYLINE_CIRCLE: [ground, [[5, 15]], S(14)], INTERSECTION_POLYLINE_POLYLINE: [[[0, -1], [0, 10]], ground],
  INTERSECTION_SEGMENT_POLYLINE: [[[0, -1, 0, 10, -10, 0, 0, 0, 10, 5]]], INTERSECTION_SPIRAL_SEGMENT: [[[1, 0.5], [3, 0.5]], [[0, 0]], S(2), S(10)],
  LATERALPILE_BROMS_COHESIONLESS: [S(5), S(30), S(18), S(0.5), S(1)], LATERALPILE_BROMS_COHESIVE: [S(100), S(1), S(50), S(0.5)],
  LATERALPILE_COHESIONLESS_REACTION_SHEAR_MOMENT: [S(5), S(30), S(18), S(0.5), S(1)], LATERALPILE_COHESIVE_REACTION_SHEAR_MOMENT: [S(100), S(1), S(50), S(0.5)],
  LOGSPIRAL: [[[0, 0]], S(2), S(30)], PILE_CAPACITY_CPT_SCHMERTMAN: [cptSoil, S(0.09), S(1.2), S(10), S("concrete")],
  PILE_SHAFT_ALPHA: [M(prof.map((r) => r[0])), M(prof.map((r) => r[1])), S(100), S(1.2), S(15), S("terzaghi")],
  PILE_SIDEFRICTION_ALPHA: [S(0.5), S("reese")], PILE_SIDEFRICTION_LAMBDA: [S(15)], PILE_TIP_MEYERHOF: [S(100), S(0), S(30)],
  PILE_TIP_MEYERHOF_FACTOR: [S(30)], PILE_TIP_VESIC: [S(150), S(100), S(0), S(30), S(0.6)], PILE_TIP_VESIC_FACTOR: [S(30), S(50)],
  PMT_ALPHA: [S("sand"), S(10)], PMT_BEARING_LIMITPRESSURE: [M(pyEpl.map((r) => r[0])), M(pyEpl.map((r) => r[2])), S(2), S(1)],
  PMT_BEARING_MENARD: [M(pyEpl.map((r) => r[0])), M(pyEpl.map((r) => r[2])), S(2), S(10), S(1), S("II"), S(0.018)], PMT_BEARING_MENARD_K: [S(1), S(0.2), S("II")],
  PMT_LAYERS: [S(2), S(1)], STEPPROFILE: [M([1, 2, 3]), M([10, 20, 30])], SPT_STEPPROFILE: [M([1, 2, 3]), M([10, 20, 30])],
  CPT_STEPPROFILE: [M([1, 2, 3]), M([10, 20, 30])], PMT_STAIRCASE: [M([1, 2, 3]), M([10, 20, 30])], POLYGON_AREA_CENTROID: [sq],
  POLYGON_INSIDE: [S(1), S(1), M([0, 2, 2, 0, 0]), M([0, 0, 2, 2, 0])], POLYLINE_SUBDIVIDE: [S(1.5), [[0, 0], [1, 1], [2, 4]]],
  SETTLEMENT_SCHMERTMAN_FACTOR: [S(2), S(0.6), S(2), S(1)], SLOPE_CIRCLE_COHESIVE: [ground, [[5, 15]], S(14), S(30), S(18)],
  SLOPE_CIRCLE_POLYLINE: [ground, [[5, 15]], S(14)], SLOPE_CIRCLE_POLYLINE_COHESIVE: [ground, [[5, 15]], S(14), S(30), S(18)],
  SLOPE_FRICTION_CIRCLE: [ground, [[5, 15]], S(14), S(10), S(20), S(18)], SLOPE_FRICTION_CIRCLE_FCT: [S(1.5), S(0.35), S(10), S(500), S(1.05), S(14), S(3), S(4), S(0.4), S(12)],
  SLOPE_FRICTION_CIRCLE_TAYLOR: [S(1.5), S(0)], SLOPE_FRICTION_FS: [S(1.5), S(0.35), S(10), S(500), S(1.05), S(14), S(3), S(4), S(0.4), S(12)],
  SLOPE_LOGSPIRAL: [ground, [[5, 15]], S(14), S(10), S(10), S(18)], SPT_DEPTHCORRECTION: [S(50)], SPT_FRICTIONANGLE: [S(20)],
  STEP_TABLE: [M([1, 2, 3]), M([10, 20, 30])], INTERSECTION_POLYGON_LINE_SLOPE: [[[0, 0], [2, 0], [2, 2], [0, 2], [0, 0]], [[1, 1]], S(0.3)],
  GRAINSIZEANALYSIS: [M([4.75, 2, 0.85, 0.425, 0.25, 0.15, 0.075]), M([95, 80, 60, 40, 25, 15, 8])],
  PILE_CAPACITY_ALPHA: [prof.map(([z, s]) => [z, s * 2]), S(100), S(0.09), S(1.2), S(15), S("terzaghi")],
  PILE_CAPACITY_CPT_LCPC: [zqc.map(([z, q]) => [z, q, "sand"]), S(0.09), S(1.2), S(15), S("concrete")],
  PILE_CAPACITY_FELLENIUS: [prof.map(([z]) => [z, 32, 18 * z, "sand"]), S(0.09), S(1.2), S(15)],
  PILE_CAPACITY_PMT_LCPC: [pyEpl.map(([z, E, pl]) => [z, pl, 18 * z / 1000, "sand"]), S(0.09), S(1.2), S(12), S("steel")],
  PILE_CAPACITY_SPT_MEYERHOF: [prof.map(([z, N]) => [z, N / 2]), S(100), S(0.09), S(1.2), S(15)],
  PILE_SHAFT_LAMBDA: [prof.map(([z, s]) => [z, s, 18 * z]), S(1.2), S(15)], PMT_SHAFTFRICTION: [S(1.5), S(3)], PMT_PILE_SHAFTFRICTION: [S(1.5), S("q3")],
  PMT_SETTLEMENT_LAMBDA: [S(2)], PMT_LAMBDA: [S(2)], PMT_MODULI: [pyEpl, S(2), S(1)],
  PMT_BEARING_EUROCODE: [M(pyEpl.map((r) => r[0])), M(pyEpl.map((r) => r[2])), S(2), S(1), S(10), S("sand gravel"), S(0.018)],
  PMT_BEARING_EUROCODE_K: [S(2), S(10), S(1), S(1.2), S("clay silt")], PMT_SETTLEMENT_BASIC: [S(0.2), S(10), S(12), S(2), S(2), S(0.5), S(0.5)],
  PMT_SETTLEMENT: [S(0.2), pyEpl, S(2), S(1), S(4), S("sand")], SHEARMOMENT: [M([0, 2, 4]), M([0, 20, 40]), S(10)], BEAM: [M([0, 2, 4]), M([0, 20, 40])],
  PRESSURE_SHEAR_MOMENT: [M([0, 2, 4]), M([0, 20, 40])], PRESSURE_SHEAR_MOMENT_ANCHOR: [M([0, 2, 4, 6]), M([0, 20, 40, 60]), S(1), S(-30)],
  SPT_BEARING_DIRECT: [S(15), S(2), S(1), S("meyerhof")], CLASSIFY_ASTM2487: [S(60), S(20), null, null, S(24), S(18)],
};
const missing = Object.keys(F).filter((k) => !(k in calls)), errs = [];
for (const [name, args] of Object.entries(calls)) {
  let res;
  try { res = F[name](...args); } catch (e) { errs.push([name, "THREW " + e.message]); continue; }
  if (!Array.isArray(res) || !Array.isArray(res[0])) { errs.push([name, "not a 2-D result"]); continue; }
  const w = res[0].length;
  if (res.some((row) => row.length !== w)) errs.push([name, "ragged result"]);
  const bad = res.flat().filter(isErr);
  const intendedNA = name === "VECTOR_LINE" || name === "CPT_ROBERTSON_LAYERS" || name === "SLOPE_FRICTION_CIRCLE";   // separators / padding / no root for this test slope (needs a reference case)      // #N/A separators / HSTACK padding
  if (bad.length && !intendedNA) errs.push([name, bad.map((b) => b.xlError).join(",") + " in " + res.length + "x" + w]);
}
console.log(`textbook checks: ${pass} passed, ${fail} failed`);
console.log(`smoke run: ${Object.keys(calls).length} functions called, ${errs.length} with errors; not covered: ${missing.join(", ") || "none"}`);
for (const e of errs) console.log("  ", e[0], "->", e[1]);
process.exitCode = fail || errs.length ? 1 : 0;
