"""Updater test bed, step 2: check the updated copies written by run.js.

Usage: python tests/updater/check.py        (after: node tests/updater/run.js)
Writes tests/updater/REPORT.md and exits with status 1 if any workbook fails.

Checks for every test workbook (original -> updated copy):
  1. The update succeeded and the copy is a valid ZIP/XLSX.
  2. Every library definition is present as a workbook-level name.
  3. No dotted geoXL/GeoEXCEL LAMBDA name is left.
  4. Every cell formula is unchanged except for dotted -> underscore renames, removal of the old
     GeoEXCEL. / GEOXL. / _xldudf_GEOXL_ prefixes, and the two
     older-argument renames (SPT_bearing with 4 arguments, PMT_settlement with 8).
  5. Every call to a geoXL function has a valid number of arguments for the new library.
  6. No cell calls an unknown function (it would show #NAME? in Excel).
  7. Updating the updated copy again changes nothing (idempotent).
Excel itself is not run: numerical results are checked separately in Excel.
"""
import json, re, sys, zipfile, html
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent.parent
import os
OUT = Path(os.environ.get("TESTBED_OUT", HERE / "out"))
REPORT = Path(os.environ.get("TESTBED_REPORT", HERE / "REPORT.md"))
run = json.loads((OUT / "run.json").read_text())
lib = run["library"]

# ---------------------------------------------------------------- library signatures
def lambda_params(compiled):
    """compiled is the Excel-stored text, e.g. _xlfn.LAMBDA(_xlpm.a,_xlop.b,body)."""
    m = re.match(r"_xlfn\.LAMBDA\(", compiled)
    if not m: return None
    args, depth, cur, i, s = [], 0, "", m.end(), compiled
    while i < len(s):
        c = s[i]
        if c == '"':
            j = i + 1
            while j < len(s):
                if s[j] == '"' and s[j + 1:j + 2] == '"': j += 2; continue
                if s[j] == '"': break
                j += 1
            cur += s[i:j + 1]; i = j + 1; continue
        if c in "({": depth += 1
        if c in ")}":
            if depth == 0: args.append(cur); break
            depth -= 1
        if c == "," and depth == 0: args.append(cur); cur = ""
        else: cur += c
        i += 1
    params = args[:-1]
    req = sum(1 for p in params if not p.startswith("_xlop."))
    return req, len(params)

SIG = {}
for d in lib["defs"]:
    SIG[d["name"].lower()] = lambda_params(d["compiled"]) if d["lambda"] else None
LIBNAMES = {d["name"].lower() for d in lib["defs"]}
CANON = {d["name"].lower(): d["name"] for d in lib["defs"]}

# ---------------------------------------------------------------- workbook reading
NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
def unx(s): return html.unescape(s.replace("\r\n", "\n").replace("\r", "\n"))
def read_book(path):
    z = zipfile.ZipFile(path)
    bad = z.testzip()
    wb = z.read("xl/workbook.xml").decode("utf-8")
    names = []
    for m in re.finditer(r"<(?:\w+:)?definedName\b([^>]*?)(?:/>|>([\s\S]*?)</(?:\w+:)?definedName>)", wb):
        attrs, body = m.group(1), m.group(2) or ""
        nm = re.search(r'\bname="([^"]*)"', attrs).group(1)
        loc = re.search(r'\blocalSheetId="([^"]*)"', attrs)
        names.append((unx(nm), loc.group(1) if loc else None, unx(body)))
    formulas = {}
    for part in sorted(n for n in z.namelist() if re.match(r"xl/worksheets/[^/]+\.xml$", n)):
        x = z.read(part).decode("utf-8")
        # cells are either self-closing (<c .../>) or <c ...>…</c>; only the latter can hold a formula
        for c in re.finditer(r'<(?:\w+:)?c\b([^>]*?)(?:/>|>(.*?)</(?:\w+:)?c>)', x, re.S):
            body = c.group(2)
            if not body or "<f" not in body and ":f" not in body: continue
            ref = re.search(r'\br="([A-Z]+\d+)"', c.group(1))
            f = re.search(r"<(?:\w+:)?f\b[^>/]*>(.*?)</(?:\w+:)?f>", body, re.S)
            if ref and f and f.group(1): formulas[(part, ref.group(1))] = unx(f.group(1))
    return {"bad": bad, "names": names, "formulas": formulas, "parts": {n: z.read(n) for n in z.namelist() if not n.endswith("/")}}

# ---------------------------------------------------------------- formula analysis
TOK = re.compile(r'"(?:[^"]|"")*"|\'(?:[^\']|\'\')*\'|[A-Za-z_\\][A-Za-z0-9_.\\]*|\d+(?:\.\d*)?(?:[Ee][+-]?\d+)?|\S')
def tokens(f):
    out, i = [], 0
    while i < len(f):
        if f[i] == "[":                      # bracket group, possibly nested: [[#Headers],[5]]
            depth, j = 0, i
            while j < len(f):
                if f[j] == "'" and j + 1 < len(f): j += 2; continue
                if f[j] == "[": depth += 1
                elif f[j] == "]":
                    depth -= 1
                    if depth == 0: break
                j += 1
            out.append(f[i:j + 1]); i = j + 1; continue
        if f[i].isspace(): i += 1; continue
        m = TOK.match(f, i); out.append(m.group(0)); i = m.end()
    return out

def calls(f):
    """Yield (name, argcount) for every function call in formula f."""
    t = tokens(f); stack = []
    for i, v in enumerate(t):
        if v == "(":
            name = t[i - 1] if i and re.match(r"[A-Za-z_\\]", t[i - 1]) else None
            if name and i >= 2 and t[i - 2] == "!": name = "!" + name   # qualified: external workbook or sheet
            stack.append([name, 0, 0, i])            # name, commas, braces, start
        elif v == "{" and stack: stack[-1][2] += 1
        elif v == "}" and stack: stack[-1][2] -= 1
        elif v == "," and stack and not stack[-1][2]: stack[-1][1] += 1
        elif v == ")" and stack:
            name, commas, _, start = stack.pop()
            if name:
                empty = (i == start + 1)
                yield name, 0 if empty else commas + 1

def local_names(f):
    """Names bound by LET/LAMBDA inside a cell formula (stored with _xlpm.)."""
    return {m.lower() for m in re.findall(r"_xlpm\.([A-Za-z_][\w.]*)", f)}

def is_builtin(n):
    return bool(re.match(r"_xl(fn|ws|udf)\.", n, re.I)) or bool(re.fullmatch(r"[A-Z][A-Z0-9.]*", n))

# Older argument lists that continue under another current name (same rule as the updater, written independently)
RENAMED = {"spt.bearing": "SPT_bearing_direct", "pmt.settlement": "PMT_settlement_basic"}
ARG_RULES = {"spt_bearing": {4: "SPT_bearing_direct"}, "pmt_settlement": {8: "PMT_settlement_basic"}}
def apply_arg_rules(toks):
    toks = list(toks); stack = []
    for i, v in enumerate(toks):
        if v == "(": stack.append([i, 0, 0])
        elif v == "{" and stack: stack[-1][2] += 1
        elif v == "}" and stack: stack[-1][2] -= 1
        elif v == "," and stack and not stack[-1][2]: stack[-1][1] += 1
        elif v == ")" and stack:
            start, commas, _ = stack.pop()
            if start and not (start >= 2 and toks[start - 2] == "!"):
                rule = ARG_RULES.get(toks[start - 1].lower().replace(".", "_"))
                n = 0 if i == start + 1 else commas + 1
                if rule and n in rule: toks[start - 1] = rule[n]
    return toks

PREFIX = re.compile(r"^(?:_xldudf_(?:geoxl|geoexcel)_|(?:geoxl|geoexcel)\.)(.+)$", re.I)
def rename(f, mapping):
    out = []
    toks = tokens(f)
    for idx, v in enumerate(toks):
        if re.match(r"[A-Za-z_]", v):
            k = v.lower()
            m = PREFIX.match(k)
            nxt = toks[idx + 1] if idx + 1 < len(toks) else ""
            if m and nxt == "(" and m.group(1).replace(".", "_") in CANON:
                if out and out[-1] == "!":                     # drop an add-in/workbook qualifier
                    out.pop(); out.pop()
                    if out and out[-1].startswith("["): out.pop()
                out.append(CANON[m.group(1).replace(".", "_")]); continue
            if k in mapping: v = mapping[k]
            elif "." in k and k.replace(".", "_") in CANON: v = CANON[k.replace(".", "_")]   # dotted call to a library function
        out.append(v)
    return "".join(apply_arg_rules(out))

# ---------------------------------------------------------------- run
rows, failures, details = [], 0, []
for r in run["results"]:
    name, problems, notes = r["file"], [], []
    if not r["ok"]:
        rows.append((r["set"], name, "FAIL", "", "", "", "", "", "")); failures += 1
        details.append(f"### {name}\n- Update failed: {r['error']}\n"); continue
    orig = read_book(ROOT / r["source"]); new = read_book(OUT / name); again = read_book(OUT / name.replace(".xlsx", ".second.xlsx"))
    if new["bad"] or again["bad"]: problems.append("corrupt ZIP entry")
    gnames = {n.lower(): body for n, loc, body in new["names"] if loc is None}
    allnames = {n.lower() for n, _, _ in new["names"]}
    miss = sorted(LIBNAMES - set(gnames))
    if miss: problems.append("library names missing: " + ", ".join(miss))
    dotted_left = [n for n, _, body in new["names"] if "." in n and not n.startswith(("_xlnm", "__xl")) and "LAMBDA" in body]
    if dotted_left: problems.append("dotted LAMBDA names left: " + ", ".join(dotted_left))
    # dotted -> underscore mapping from the original
    mapping = {}
    for n, _, body in orig["names"]:
        if "." in n and not n.startswith(("_xlnm", "__xl")) and "LAMBDA" in body:
            mapping[n.lower()] = next((x for x, _, _ in new["names"] if x.lower() == n.lower().replace(".", "_")), n.replace(".", "_"))
            alt = RENAMED.get(n.lower())
            if alt and SIG.get(alt.lower()) and lambda_params(body) and lambda_params(body)[1] == SIG[alt.lower()][1]:
                mapping[n.lower()] = alt      # older argument list: continues under the other current name
    # 4. formulas unchanged except renames
    changed = renamed = 0
    for key, f0 in orig["formulas"].items():
        f1 = new["formulas"].get(key)
        if f1 is None: problems.append(f"formula disappeared at {key}"); continue
        if f1 != f0:
            if rename(f0, mapping) == "".join(tokens(f1)): renamed += 1
            else:
                changed += 1
                if changed <= 3: problems.append(f"formula changed at {key[0].split('/')[-1]}!{key[1]}: {f0[:80]} -> {f1[:80]}")
    if set(new["formulas"]) - set(orig["formulas"]): problems.append("new formula cells appeared")
    # 5-6. calls
    def unknown_in(book, names):
        found = set()
        for key, f in book["formulas"].items():
            loc = local_names(f)
            for fn, nargs in calls(f):
                k = fn.lower()
                if k.startswith("!") or k in names or k in loc or k.startswith("_xlpm.") or is_builtin(fn) or k in SIG: continue
                found.add(fn)
        return found
    before_unknown = unknown_in(orig, {n.lower() for n, _, _ in orig["names"]})
    external = sorted({fn[1:] for f in new["formulas"].values() for fn, _ in calls(f) if fn.startswith("!")})
    if external: notes.append("calls to functions in another workbook ([1]!name), left as they are: " + ", ".join(external))
    ncalls = unknown = arity = 0; unknown_names, arity_list = set(), []
    for key, f in new["formulas"].items():
        loc = local_names(f)
        for fn, nargs in calls(f):
            k = fn.lower()
            if k in SIG:
                ncalls += 1
                sig = SIG[k]
                if sig and not (sig[0] <= nargs <= sig[1]):
                    arity += 1; arity_list.append(f"{key[0].split('/')[-1]}!{key[1]} {fn}: {nargs} arguments, expects {sig[0]}–{sig[1]}")
            elif k.startswith("!") or k in allnames or k.replace("_xlpm.", "") in loc or k.startswith("_xlpm.") or is_builtin(fn):
                continue
            else:
                unknown += 1; unknown_names.add(fn)
    if arity:
        orig_bad = set()
        for key, f in orig["formulas"].items():
            for fn, nargs in calls(f):
                sig = SIG.get(fn.lower())
                if sig and not (sig[0] <= nargs <= sig[1]): orig_bad.add(key)
        problems.append(f"{arity} geoXL calls with a wrong number of arguments for {lib['version']}" + (" (also wrong in the original, against the same signature)" if orig_bad else "") + ":")
        problems += ["  " + a for a in arity_list[:5]]
    introduced = unknown_names - before_unknown
    if introduced: problems.append("unknown functions introduced by the update (would show #NAME?): " + ", ".join(sorted(introduced)))
    if unknown_names & before_unknown: notes.append("unknown functions already in the original: " + ", ".join(sorted(unknown_names & before_unknown)))
    # 7. idempotent
    s = r["second"]
    if s["added"] or s["removedDottedNames"] or s["renamedFormulas"]: problems.append(f"second update still changed things: {s}")
    strip = lambda names: sorted((n, l or "", b) for n, l, b in names if n not in ("_GeoEXCEL_UpdateInfo", "_GeoLabXL_UpdateInfo"))
    if strip(new["names"]) != strip(again["names"]): problems.append("second update changed the defined names")
    if new["formulas"] != again["formulas"]: problems.append("second update changed cell formulas")
    custom = sorted(n for n, _, body in new["names"] if "LAMBDA" in body and n.lower() not in LIBNAMES)
    if custom: notes.append("other LAMBDAs kept: " + ", ".join(custom))
    rep = r["report"]; c = rep["counts"]
    for w in rep["warnings"]: notes.append("updater note: " + w)
    status = "PASS" if not problems else "FAIL"
    failures += status == "FAIL"
    rows.append((r["set"], name, status, rep["previousVersion"], f"{c['updated']} / {c['added']}", c["removedDottedNames"], renamed, ncalls, "yes" if not any("second update" in p for p in problems) else "NO"))
    shown = [n for n in notes if not n.startswith("updater note: Library version changed")]
    if problems or shown:
        details.append(f"### {name}\n" + "".join(f"- **{p}**\n" for p in problems) + "".join(f"- {n}\n" for n in notes if not n.startswith("updater note: Library version changed")))

md = [f"# Updater test bed report\n",
      f"Library **{lib['version']}** ({lib['n']} definitions, SHA256 `{lib['hash'][:12]}…`) · {len(rows)} workbooks · "
      f"**{len(rows) - failures} passed, {failures} failed**\n",
      "| Set | Workbook | Result | Previous version | Updated / added | Dotted names converted | Cell formulas renamed | geoXL calls checked | Idempotent |",
      "|---|---|---|---|---|---|---|---|---|"]
md += ["| " + " | ".join(str(x) for x in row) + " |" for row in rows]
md += ["", "Checks: valid file; all library names present; no dotted LAMBDA names left; cell formulas unchanged except renames (dotted names, and SPT_bearing with 4 arguments → SPT_bearing_direct, PMT_settlement with 8 → PMT_settlement_basic); "
       "argument counts valid for the new library; no unknown functions; a second update changes nothing. "
       "Excel is not run here — numerical results are checked in Excel.", ""]
if details: md += ["## Details", ""] + details
REPORT.write_text("\n".join(md) + "\n", encoding="utf-8")
print("\n".join(md[:2])); print(f"{len(rows) - failures} passed, {failures} failed — see tests/updater/REPORT.md")
sys.exit(1 if failures else 0)
