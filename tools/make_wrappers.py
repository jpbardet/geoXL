"""Generate src/wrappers.js (Excel custom-function wrappers with JSDoc) from a compact spec table.
Descriptions are taken from the GeoEXCEL LAMBDA library so both stay identical."""
import re, sys, json

LIB = sys.argv[1]          # GeoEXCEL_v0.13.txt
OUT = sys.argv[2]          # src/wrappers.js

src = open(LIB, encoding="utf-8").read()
desc = {}
for m in re.finditer(r"/\*\*\s*(.*?)\s*\*/\s*\n\s*([A-Za-z_][\w]*)\s*=\s*LAMBDA", src, re.S):
    desc[m.group(2)] = " ".join(m.group(1).split())

# kind: L = element-by-element (scalar result), R = one row per element, F = whole result
# param codes: n number, s text, on/os optional (lifted), v column, m table, p point {x,y},
#              P two-column polygon/polyline, O optional (not lifted)
SPEC = """
version|F|
bearing_capacity|L|coh n, phi n, q n, gam n, B n, L n, D n, H n, V n, author s, A_ os
bearing_depthFactors|F|DB n, phi n, author s
bearing_factors|F|phi n, author_ s
bearing_inclinationFactors|F|H n, V n, BL n, phi n, coh n, author s, A_ O
bearing_Nq|L|Phi n, author_ s
bearing_shapeFactors|F|BL n, phi n, author s, opt_ O
bearing_Skempton|L|DB n, BL n
classify_AASHTO|L|pf10_ n, pf40_ n, pf200_ n, LL_ os, PL_ os
classify_ASTM|L|list m, pf4_ n, pf200_ n, Cu_ os, Cc_ os, LL_ os, PL_ os, organic_ os
classify_Plasticity|L|LL n, PL n, organic os
classify_USC|L|pf4_ n, pf200_ n, Cu_ os, Cc_ os, LL_ os, PL_ os, organic_ os
consolidation_test|F|Tarray v, Xarray v, ta n, tb n, hi n
consolidation_Tv|L|U n
consolidation_U|F|tv v, case_ O
consolidation_Ux|F|Xarray v, tv n, case_ O
consolidation_U_approximate|L|tv n
consolidation_U_variable|L|tv n, tc n
cosd|L|X n
Coulomb_cohesiveforces|F|t n, phi n, alpha n, beta n, delta n, g n, h n, coh_ O
Coulomb_forces|F|t n, phi n, alpha n, beta n, delta n, g n, h n
coulomb_K|L|phi n, alpha n, beta n, delta n, passive os
Coulomb_maximum|F|a_ n, b_ n, phi n, alpha n, beta n, delta n, g n, h n, iter_ O
coulomb_theta|L|phi n, alpha n, beta n, delta n
cutline|F|Xarray v, Yarray v, slope n, intercept n
elastic_circular_radius|L|s n
elastic_circular_stress|R|rValue n, z n, nu_ os
elastic_embankment|L|b1z n, b2z n
elastic_Janbu_mu0|L|DB n, LB n
elastic_Janbu_mu1|L|LB n, DB s
elastic_lineload|R|x n, y n
elastic_lineload_horizontal|R|x n, y n
elastic_Newmark_factor|L|m n, n n
elastic_Newmark_stress|L|x_ n, y_ n, z_ n, pt P
elastic_pointload|R|rValue n, z n, nu_ os
elastic_sector|L|x n, y n, z n, pts P
elastic_sector_row|F|txy v
elastic_sector_slice|L|ra n, rb n, z n, dt n
elastic_stripload|R|B n, x n, y n
elastic_triangle|R|B n, x n, y n
Interpolate|L|X n, Xa v, Ya v, mode O, outside O
intersection_polygon_line|F|polygon P, pt p, t n
intersection_polygon_ray|F|polygon P, pt p, t n
minimum|L|A n, B n, iter_ os
polygon_centroidArea|F|polygon P
polygon_weightMoment|F|XY P, g n, Pt p
pressure_forceMoment|F|d v, p v, d0 n
Rankine_K|L|phi n, beta n
Rankine_stress|F|phi n, coh n, alpha n, beta n, gz n, passive O
sind|L|X n
SPT_bearing|F|z v, N v, B n, D n, author s
SPT_BurlandBurbidge_settlement|F|z v, N v, q n, B n, LB O, D_ O, H O, sc O
SPT_correction_depth|L|s n
SPT_Correlation_friction|L|N n, author s, s os, pa os
stress_ForceMoment|F|A p, B p, sa p, sb p
Stress_Mohr|R|sa n, ta n, alpha n, sm n
stress_Mohr_failure|F|sa n, ta n, A n, phi n, coh n, passive O
stress_Mohr_pole|F|sa n, ta n, A n, sb n
stress_Mohrcircle|F|sm n, rValue n
tand|L|X n
vector_line|F|ptA p, ptB p, sigA p, sigB p, M n, s n
water_pressure|L|depth n, waterTableDepth n, unitWeight os
elastic_stripload_xy|R|B n, X n, Y n
circle|F|pt p, rValue n, ta_ O, tb_ O, n_ O
classify_CPT_LCPC|R|qc n, soil s, pile s
classify_PMT_LCPC|R|pl n, soil s, pile s
CPT_average_qc|F|z v, qc v, D n, B n, author s
CPT_bearing_Meyerhof_k|L|B n, D n, dw n, m os
CPT_bearing_Robertson_k|L|B n, D n, shape s, density s
CPT_friction_angle|L|z n, qc n, g n, author_ s
CPT_Olsen_layers|F|cpt m
CPT_Olsen_SCN|F|dqs m, pa n, n n, g n, gw n, dw n
CPT_Olsen_SCN_row|L|Qt n, Fr n
CPT_Robertson_classification_chart|L|x_ n, sel s
CPT_Robertson_layers|F|cpt m
CPT_Robertson_SBT|F|dqs m, pa n, n n, g n, gw n, dw n
CPT_Robertson_SBT_row|L|Qt_ os, Fr_ os
CPT_settlement|F|zqc m, k n, B n, D n, L n, qn n, g n
CPT_settlement_Iz|L|z n, B n, D n, L n, qn n, g n
distance_point_polyline|F|M p, pl P
interpolate_table|L|x n, Y n, XY m
intersection_polyline_circle|F|pline P, center p, radius n
intersection_polyline_polyline|F|pl P, ql P, m O
intersection_segment_polyline|F|row v
intersection_spiral_segment|F|AB_ P, pt p, rValue n, phi n, x_ O, iter_ O
lateralpile_Broms_cohesionless|F|D n, phi n, g n, B n, e n
lateralpile_Broms_cohesive|F|V n, e n, Su n, B n
lateralpile_cohesionless_reaction_shear_moment|F|D n, phi n, g n, B n, e n
lateralpile_cohesive_reaction_shear_moment|F|V n, e n, Su n, B n, D_ O, Dr_ O
logSpiral|F|pt p, rValue n, phi n, ta_ O, tb_ O, n_ O
pile_capacity_CPT_Schmertman|F|z_qc_fs_soil m, A n, p n, L n, pile s
pile_shaft_alpha|F|z v, Su v, pa n, p n, L n, author s
pile_sidefriction_alpha|L|Su n, author s
pile_sidefriction_lambda|L|D_ n, ft os
pile_tip_Meyerhof|F|pa n, cValue n, phi n
pile_tip_meyerhof_factor|F|phi n
pile_tip_vesic|F|sv n, pa n, cValue n, phi_ O, Dr_ O
pile_tip_vesic_factor|F|phi n, Ir n
PMT_alpha|L|soil s, Epl n
PMT_bearing_limitPressure|F|z v, pl v, B n, D n
PMT_bearing_Menard|F|z v, pl v, B n, D n, L n, soil s, g n
PMT_bearing_Menard_k|L|H n, BL n, soil s
PMT_layers|F|B n, D n
stepProfile|F|x v, y v
SPT_stepProfile|F|x v, y v
CPT_stepProfile|F|x v, y v
PMT_staircase|F|x v, y v
polygon_area_centroid|F|polygon P
Polygon_inside|L|x n, y n, xa v, ya v
polyline_subdivide|L|X n, polyline P
settlement_Schmertman_factor|L|z n, Imax n, B n, za n, BL_ os
slope_circle_cohesive|F|pline P, center p, radius n, cu n, unitweight n, detail O
slope_circle_polyline|F|pl P, center p, radius n
slope_circle_polyline_cohesive|F|pline P, center p, radius n, cu n, unitweight n, detail O
slope_friction_circle|F|pline P, pt p, rValue n, coh n, phi n, g n, opt_ O
slope_friction_circle_fct|L|FS n, f n, coh n, w n, kr n, rValue n, OD n, DM n, alpha n, Lc n
slope_friction_circle_Taylor|L|t n, opt n
slope_friction_FS|L|FS n, f n, coh n, w n, kr n, rValue n, a n, DM n, alpha n, ab n
slope_logSpiral|F|pl P, pt p, rValue n, phi n, coh n, g n
SPT_depthCorrection|L|s n
SPT_frictionAngle|L|N n, spa_ os, author_ os
Step_table|F|x v, y v
intersection_polygon_line_slope|F|polyline P, M p, t n
grainsizeanalysis|F|d v, p v
pile_capacity_alpha|F|zSu m, pa n, a n, p n, L n, author s, D_ O
pile_capacity_CPT_LCPC|F|z_qc_soil m, a n, p n, L n, pile s
pile_capacity_Fellenius|F|z_phi_sv_soil m, a n, p n, L n, D_ O
pile_capacity_PMT_LCPC|F|z_pl_sv_soil m, a n, p n, L n, pile s
pile_capacity_SPT_Meyerhof|F|zN m, pa n, a n, p n, L n, disp O
pile_shaft_lambda|F|z_Su_sv m, p n, L n, ft O
PMT_shaftFriction|L|pl n, i n
PMT_pile_shaftFriction|L|pl n, classi s
PMT_settlement_lambda|F|LB n, opt_ O
PMT_lambda|F|LB n
PMT_moduli|F|yEpl m, B n, D n
PMT_bearing_Eurocode|F|z v, pl v, B n, D n, L n, soil s, g n
PMT_bearing_eurocode_k|L|B n, L n, D n, pl n, soil_ s
PMT_settlement_basic|F|q n, Ec n, Ed n, B n, LB n, ac n, ad n, B0_ O
PMT_settlement|F|q n, yEpl m, B n, D n, L n, soil s, ft_ O
shearMoment|F|d v, p v, k n
beam|F|d v, p v, n O
pressure_shear_moment|F|d v, p v, n O
pressure_shear_moment_anchor|F|d v, p v, da n, Ta n, n O
SPT_bearing_direct|L|N n, B n, D n, author s
classify_ASTM2487|R|pf4 n, pf200 n, cu_ os, cc_ os, ll_ os, pl_ os, ratio_ os
"""

# calls that differ from G.<name>(converted params)
CALL = {
    "version": "G.version()",
    "polygon_centroidArea": "G.polygon_area_centroid({polygon})",
    "SPT_stepProfile": "G.stepProfile({x}, {y})",
    "CPT_stepProfile": "G.stepProfile({x}, {y})",
    "PMT_staircase": "G.stepProfile({x}, {y})",
    "SPT_depthCorrection": "G.SPT_correction_depth({s})",
    "PMT_lambda": 'G.PMT_settlement_lambda({LB}, "a")',
    "beam": "G.shearMoment({d}, {p}, {n} === undefined ? 20 : num({n}))",
    "pressure_shear_moment": "[SM_HDR, ...G.shearMoment({d}, {p}, {n} === undefined ? 20 : num({n}))]",
    "elastic_sector_row": "(() => {{ const v = {txy}, poly = []; for (let i = 1; i + 1 < v.length; i += 2) poly.push([v[i], v[i + 1]]); return G.elastic_sector_rowValue(v[0], poly); }})()",
    "Interpolate": "G.interp({X}, {Xa}, {Ya}, {mode}, {outside})",
}

TYPEDOC = {"n": "any[][]", "s": "any[][]", "on": "any[][]", "os": "any[][]", "v": "any[][]", "m": "any[][]", "p": "any[][]", "P": "any[][]", "O": "any[][]"}
HINT = {"n": "", "s": " (text)", "os": " (optional)", "on": " (optional)", "v": " (column of values)", "m": " (table)",
        "p": " (point {x, y})", "P": " (two-column x, y list)", "O": " (optional)"}
PARAM_TEXT = {
    "phi": "Friction angle (degrees)", "author": "Method name", "author_": "Method name", "B": "Width", "L": "Length", "D": "Depth",
    "g": "Unit weight", "z": "Depth", "coh": "Cohesion", "pa": "Atmospheric pressure", "alpha": "Angle (degrees)", "beta": "Angle (degrees)",
    "delta": "Wall friction angle (degrees)", "h": "Height", "pile": "Pile type", "soil": "Soil type",
}

def esc(t): return t.replace("*/", "* /")

lines = ["/* ============================================================================================",
         " * Part 3: Excel custom functions (generated by tools/make_wrappers.py — edit the spec, not this file).",
         " * Every argument is declared any[][] so a column passed where a single value is expected",
         " * spills a column of results, as the LAMBDA versions do.",
         " * ==========================================================================================*/", ""]
names = []
for line in SPEC.strip().splitlines():
    name, kind, plist = line.split("|")
    params = [p.strip().split() for p in plist.split(",") if p.strip()]
    names.append(name)
    d = desc.get(name)
    if d is None: raise SystemExit("no description for " + name)
    doc = ["/**", " * " + esc(d), " * @customfunction " + name.upper()]
    for pn, pc in params:
        opt = pc in ("os", "on", "O")
        label = PARAM_TEXT.get(pn, pn) + HINT[pc]
        doc.append(f" * @param {{{TYPEDOC[pc]}}} {'[' + pn + ']' if opt else pn} {esc(label)}")
    doc.append(" * @returns {any[][]} Result (single value, row or table).")
    doc.append(" */")
    args = [pn for pn, _ in params]
    lines += doc
    lines.append(f"function {name}({', '.join(args)}) {{")
    # conversions
    fixed_conv = {"v": "colNum({})", "m": "mat({})", "p": "point({})", "P": "pts2({})", "O": "opt({})"}
    if kind == "F":
        conv = {}
        for pn, pc in params:
            conv[pn] = {"n": "num({})", "s": "str({})", "os": "opt({})", "on": "opt({})"}.get(pc, fixed_conv.get(pc, "{}")).format(pn)
        call = CALL.get(name, f"G.{name}(" + ", ".join("{" + a + "}" for a in args) + ")").format(**conv)
        lines.append(f"  return safe(() => {call});")
    else:
        # lifted: scalar params broadcast; table/vector/optional-fixed params converted once
        fixed_idx = [i for i, (_, pc) in enumerate(params) if pc in ("v", "m", "p", "P", "O")]
        pre = []
        lift_args = []
        for i, (pn, pc) in enumerate(params):
            if pc in fixed_conv: pre.append(f"  const {pn}_c = {fixed_conv[pc].format(pn)};"); lift_args.append(pn + "_c")
            else: lift_args.append(pn)
        inner = {}
        for pn, pc in params:
            inner[pn] = {"n": f"num(e_{pn})", "s": f"e_{pn}", "os": f"e_{pn}", "on": f"e_{pn}"}.get(pc, f"e_{pn}")
        call = CALL.get(name, f"G.{name}(" + ", ".join("{" + a + "}" for a in args) + ")").format(**inner)
        lines += pre
        fn = "lift" if kind == "L" else "liftRows"
        lines.append(f"  return safeLifted(() => {fn}((" + ", ".join("e_" + a for a in args) + f") => {call}, [{', '.join(lift_args)}], {json.dumps(fixed_idx)}));")
    lines.append("}")
    lines.append("")

assert len(set(n.upper() for n in names)) == len(names), "duplicate names"
lines.append("// ---------------------------------------------------------------- registration")
lines.append("const GEOXL_FUNCTIONS = {" + ", ".join(f'{n.upper()}: {n}' for n in names) + "};")
lines.append('if (typeof CustomFunctions !== "undefined") { for (const [id, f] of Object.entries(GEOXL_FUNCTIONS)) CustomFunctions.associate(id, f); }')
lines.append('if (typeof module !== "undefined") module.exports = { G, GEOXL_FUNCTIONS };')
open(OUT, "w", encoding="utf-8").write("\n".join(lines) + "\n")
print(len(names), "wrappers written; descriptions found for all")
