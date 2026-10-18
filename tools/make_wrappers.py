"""Generate src/wrappers.js (Excel custom-function wrappers with JSDoc) and docs/index.html (help page)
from a compact spec table. Function descriptions come from the GeoEXCEL LAMBDA library; argument help,
returns and notes come from tools/help.txt."""
import re, sys, json, os, html

LIB = sys.argv[1]                                          # library/GeoEXCEL_v0.13.txt
OUT = sys.argv[2]                                          # src/wrappers.js
HELP = sys.argv[3] if len(sys.argv) > 3 else "tools/help.txt"
PAGE = sys.argv[4] if len(sys.argv) > 4 else "docs/index.html"
HELP_URL = "https://jpbardet.github.io/geoXL/index.html#"  # Excel's "Help on this function" link

src = open(LIB, encoding="utf-8").read()
desc = {}
for m in re.finditer(r"/\*\*((?:(?!\*/).)*?)\*/\s*\n\s*([A-Za-z_][\w]*)\s*=\s*LAMBDA", src, re.S):
    desc[m.group(2)] = " ".join(m.group(1).split())

# help.txt: [name] blocks with "argument: text", "returns: text", "notes: text" lines
helpdb, cur = {}, None
for ln in open(HELP, encoding="utf-8"):
    ln = ln.rstrip("\n")
    if not ln.strip() or ln.lstrip().startswith("#"): continue
    m = re.match(r"\s*\[(\w+)\]\s*$", ln)
    if m: cur = helpdb.setdefault(m.group(1), {"params": {}, "returns": "", "notes": []}); continue
    if cur is None or ":" not in ln: raise SystemExit("help.txt: cannot read line: " + ln)
    k, v = (t.strip() for t in ln.split(":", 1))
    if k == "returns": cur["returns"] = v
    elif k == "notes": cur["notes"].append(v)
    else: cur["params"][k] = v

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
astm_table|F|
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
SPT_bearing|L|z v, N v, B n, D n, author s
SPT_BurlandBurbidge_settlement|L|z v, N v, q n, B n, LB os, D_ os, H os, sc os
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
CPT_average_qc|L|z v, qc v, D n, B n, author s
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
pages = []
for line in SPEC.strip().splitlines():
    name, kind, plist = line.split("|")
    params = [p.strip().split() for p in plist.split(",") if p.strip()]
    names.append(name)
    d = desc.get(name)
    if d is None: raise SystemExit("no description for " + name)
    doc = ["/**", " * " + esc(d), " * @customfunction " + name.upper(), " * @helpurl " + HELP_URL + name.upper()]
    pages.append((name, params, d))
    for pn, pc in params:
        opt = pc in ("os", "on", "O")
        htext = helpdb.get(name, {}).get("params", {}).get(pn)
        label = htext if htext else PARAM_TEXT.get(pn, pn) + HINT[pc]
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

# ---------------------------------------------------------------- help page (docs/index.html)
GROUPS = [("Bearing capacity", ("bearing_", "spt_bearing", "cpt_bearing", "cpt_average", "pmt_bearing")),
          ("Soil classification", ("classify_", "astm_table", "grainsize")),
          ("Consolidation", ("consolidation_", "cutline")),
          ("Earth pressure & walls", ("coulomb", "rankine", "pressure_", "shearmoment", "beam", "water_pressure")),
          ("Elastic stresses", ("elastic_",)),
          ("Settlement", ("settlement_", "spt_burland", "cpt_settlement", "pmt_settlement", "pmt_lambda", "pmt_moduli", "pmt_alpha")),
          ("Piles", ("pile_", "lateralpile_", "pmt_pile", "pmt_shaft")),
          ("In-situ tests (SPT, CPT, PMT)", ("spt_", "cpt_", "pmt_")),
          ("Slope stability", ("slope_",)),
          ("Stresses & Mohr circle", ("stress_",)),
          ("Geometry & plotting", ("polygon", "polyline", "intersection_", "distance_", "circle", "logspiral", "vector_line", "stepprofile", "step_table")),
          ("Math & utilities", ("",))]

def group_of(n):
    l = n.lower()
    for g, pre in GROUPS:
        if any(l.startswith(p) for p in pre): return g
    return GROUPS[-1][0]

E = lambda t: html.escape(t, quote=False)
cards = {g: [] for g, _ in GROUPS}
for name, params, d in pages:
    h = helpdb.get(name, {"params": {}, "returns": "", "notes": []})
    sig = ", ".join(("[" + pn + "]") if pc in ("os", "on", "O") else pn for pn, pc in params)
    rows = "".join(f"<tr><td><code>{E(pn)}</code>{' <span class=opt>optional</span>' if pc in ('os','on','O') else ''}</td><td>{E(h['params'].get(pn, PARAM_TEXT.get(pn, pn)))}</td></tr>" for pn, pc in params)
    body = f"<p class=desc>{E(d)}</p>"
    if rows: body += f"<table class=args><thead><tr><th>Argument</th><th>Meaning</th></tr></thead><tbody>{rows}</tbody></table>"
    if h["returns"]: body += f"<p><b>Returns:</b> {E(h['returns'])}</p>"
    for nt in h["notes"]: body += f"<p class=note>{E(nt)}</p>"
    cards[group_of(name)].append((name.upper(), f'<section class=fn id="{name.upper()}"><h3><a href="#{name.upper()}">=GEOXL.{name.upper()}</a><span class=sig>({E(sig)})</span></h3>{body}</section>'))

nav, main = [], []
for g, _ in GROUPS:
    if not cards[g]: continue
    gid = re.sub(r"[^a-z]+", "-", g.lower()).strip("-")
    nav.append(f'<a href="#{gid}">{E(g)} <span>{len(cards[g])}</span></a>')
    main.append(f'<h2 id="{gid}">{E(g)}</h2>' + "".join(c for _, c in sorted(cards[g])))

page = f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>geoXL functions</title>
<style>
:root{{--bg:#fff;--fg:#1d1d1f;--muted:#5f6368;--line:#e3e3e3;--accent:#1f4e79;--note:#fff6e0;--noteline:#e0b44c;--card:#fafafa}}
@media (prefers-color-scheme: dark){{:root{{--bg:#16181b;--fg:#e8e8e8;--muted:#a0a4aa;--line:#2c3036;--accent:#7fb2e5;--note:#2e2716;--noteline:#9a7a2e;--card:#1c1f23}}}}
*{{box-sizing:border-box}}
body{{margin:0;background:var(--bg);color:var(--fg);font:15px/1.5 system-ui,-apple-system,Segoe UI,sans-serif}}
.wrap{{max-width:1180px;margin:0 auto;padding:24px 16px;display:grid;grid-template-columns:230px 1fr;gap:28px}}
@media (max-width:820px){{.wrap{{grid-template-columns:1fr}} nav{{position:static!important;max-height:none!important}}}}
header{{grid-column:1/-1}} h1{{font-size:26px;margin:0 0 4px}} header p{{color:var(--muted);margin:0 0 12px}}
input{{width:100%;padding:9px 12px;border:1px solid var(--line);border-radius:6px;background:var(--bg);color:var(--fg);font-size:15px}}
nav{{position:sticky;top:16px;align-self:start;max-height:calc(100vh - 32px);overflow:auto;font-size:14px}}
nav a{{display:flex;justify-content:space-between;padding:5px 8px;border-radius:5px;color:var(--fg);text-decoration:none}}
nav a:hover{{background:var(--card)}} nav span{{color:var(--muted)}}
h2{{font-size:19px;margin:28px 0 8px;padding-bottom:4px;border-bottom:2px solid var(--line)}} h2:first-child{{margin-top:0}}
.fn{{border:1px solid var(--line);border-radius:8px;padding:12px 14px;margin:10px 0;background:var(--card);scroll-margin-top:12px}}
.fn:target{{outline:2px solid var(--accent)}}
h3{{margin:0 0 4px;font-size:15px;word-break:break-word}} h3 a{{color:var(--accent);text-decoration:none}}
.sig{{color:var(--muted);font-weight:400;font-family:ui-monospace,Menlo,monospace;font-size:13px}}
.desc{{margin:4px 0 8px}} p{{margin:6px 0}}
.args{{width:100%;border-collapse:collapse;font-size:14px}} .args th{{text-align:left;color:var(--muted);font-weight:500;font-size:12px;padding:4px 6px}}
.args td{{border-top:1px solid var(--line);padding:5px 6px;vertical-align:top}} .args td:first-child{{width:30%;white-space:nowrap}}
code{{font-family:ui-monospace,Menlo,monospace;font-size:13px}}
.opt{{font-size:11px;color:var(--muted);border:1px solid var(--line);border-radius:4px;padding:0 4px;margin-left:4px}}
.note{{background:var(--note);border-left:3px solid var(--noteline);padding:6px 10px;border-radius:4px;font-size:14px}}
.hidden{{display:none}}
</style></head><body><div class=wrap>
<header><h1>geoXL</h1>
<p>{len(pages)} geotechnical functions for Excel. Type <code>=GEOXL.</code> in a cell; Excel shows each argument's meaning as you type. Use one consistent unit system; angles are in degrees unless stated.</p>
<input id=q type=search placeholder="Filter functions (e.g. bearing, pile, Mohr, settlement)…" aria-label="Filter functions"></header>
<nav>{''.join(nav)}</nav>
<main>{''.join(main)}</main></div>
<script>
const q=document.getElementById('q');
q.addEventListener('input',()=>{{const t=q.value.trim().toLowerCase();
document.querySelectorAll('.fn').forEach(s=>s.classList.toggle('hidden',t&&!s.textContent.toLowerCase().includes(t)));
document.querySelectorAll('main h2').forEach(h=>{{let n=h.nextElementSibling,any=false;while(n&&n.tagName!=='H2'){{if(!n.classList.contains('hidden'))any=true;n=n.nextElementSibling;}}h.classList.toggle('hidden',!any);}});}});
</script></body></html>
"""
os.makedirs(os.path.dirname(PAGE) or ".", exist_ok=True)
open(PAGE, "w", encoding="utf-8").write(page)
print("help page written:", PAGE, "-", sum(len(v) for v in cards.values()), "functions")
