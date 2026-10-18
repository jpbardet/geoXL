# geoXL — geotechnical functions for Excel (JavaScript add-in)

145 Excel functions translated from the **GeoEXCEL LAMBDA library v0.13**. In Excel they appear with the
`GEOXL.` prefix, for example `=GEOXL.WATER_PRESSURE(A2, 2)` or `=GEOXL.BEARING_FACTORS(30, "vesic")`.
Function descriptions are the same as in the LAMBDA library and show up in Excel's autocomplete.

## What is in this folder

| Path | Purpose |
|---|---|
| `manifest.xml` | The file students install. It tells Excel where the code lives. |
| `docs/` | Everything GitHub Pages serves: `functions.js` (code), `functions.json` (descriptions), `functions.html` (loader), `index.html` (searchable function list), `assets/` (icons). |
| `src/` | Source: `helpers.js` (Excel-like helpers), `core.js` (the engineering functions), `wrappers.js` (generated Excel wrappers). |
| `tools/make_wrappers.py` | Generates `src/wrappers.js` from a compact table plus the LAMBDA library's descriptions. |
| `tools/help.txt` | Help for every function: the meaning of each argument (shown by Excel as you type), what it returns, and notes. The build turns it into Excel's argument tips and the help page `docs/index.html`. |
| `tests/test.js` | 45 textbook checks + a run of all 145 functions. `node tests/test.js` |
| ASTM names | The 100-row ASTM D2487 name table (formerly the `ASTMtable` name) is built in: `=GEOXL.ASTM_TABLE()`; `CLASSIFY_ASTM` uses it when its first argument is `""`. |
| `library/GeoEXCEL_v0.13.txt` | The LAMBDA source this version was translated from. |

## 1. Publish on GitHub Pages (once)

1. Create a **public** repository on GitHub named `geoxl`, and upload the contents of this folder.
2. In the repository: **Settings → Pages → Build and deployment → Deploy from a branch**, branch `main`, folder **`/docs`**. Save.
3. After a minute, `https://<your-user-name>.github.io/geoxl/` shows the function list.
4. In `manifest.xml`, replace **`YOUR-GITHUB-USERNAME`** (6 places) with your GitHub user name. Commit.

## 2. Install the manifest

**Excel for Mac**
1. Finder → Go → Go to Folder (Shift-Command-G): `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef` (create `wef` if needed).
2. Copy `manifest.xml` into it, restart Excel, open a workbook.
3. Home → Add-ins → choose **geoXL** (under Developer add-ins). Type `=GEOXL.` in a cell.

**Excel on the web**: Home → Add-ins → More Settings (or Advanced…) → **Upload My Add-in** → choose `manifest.xml`.
The upload is remembered by that browser only.

**Excel for Windows**: share a folder containing `manifest.xml`, add its network path under
File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs (tick *Show in Menu*), restart,
then Home → Add-ins → More Add-ins → Shared Folder.

**Whole class**: the university's Microsoft 365 admins can deploy the manifest centrally
(Settings → Integrated apps → Upload custom apps, using the manifest's GitHub address); students then see it under *Admin Managed*.

## 3. Updating

Edit `src/core.js` (or the spec in `tools/make_wrappers.py`), then:

```
npm install          # once, installs Microsoft's custom-functions-metadata tool
npm run build        # regenerates docs/functions.js and docs/functions.json
npm test
```

Bump `?v=1.0.0` in `manifest.xml` and `docs/functions.html` when you publish, so Excel does not keep an old cached copy.
Students who installed the manifest get the new code automatically the next time Excel loads the add-in.

## Help text

* **One-line description** of a function: the `/** … */` comment above it in `library/GeoEXCEL_v0.13.txt`.
* **Argument meanings, returns, notes**: the function's `[name]` block in `tools/help.txt`, one `argument: meaning` line per argument.
* Excel shows the argument meaning while you type the formula, and the *Insert Function* dialog links to the function's section
  of `https://jpbardet.github.io/geoXL/index.html`. `npm run build` regenerates everything (GitHub does it for you if the build workflow is set up).

## Differences from the LAMBDA version

* **Names**: `GeoEXCEL.water_pressure` → `GEOXL.WATER_PRESSURE` (Excel shows custom function names in capitals).
* **Columns**: every function that returns one value accepts columns for its single-value arguments and spills a
  column, like the LAMBDAs do. Functions that return a labelled row (e.g. `ELASTIC_POINTLOAD`) return one row per input.
* **Errors**: where a LAMBDA gives `#CALC!` (an empty `FILTER`), the add-in gives `#N/A`. Invalid arithmetic gives `#NUM!`.
* **Workbooks need the add-in**: a workbook using `GEOXL.` functions shows `#NAME?` on a computer without the add-in.

## Behaviour kept from the LAMBDAs (worth reviewing in both versions)

* `interpolate_table`: the fourth corner uses `INDEX(Za, r+1, c)`, the same cell as another corner; it was probably meant to be `INDEX(Za, r, c+1)`.
* `classify_USC`: the branch `fines = "CL-ML"` can never be true, because `classify_Plasticity` returns `"CL & ML"`.
* `consolidation_test`: `tc` and `xc` both come from the y value of the first intersection, so `Cv = 0.848 / xc²`.
* `elastic_sector`: 361 rays (0° and 360° both counted) divided by 360.
* `Step_table`: the LAMBDA uses `INDEX(y, 0)` for the first row; the add-in returns the first y value there.
* `classify_ASTM`: the row choice for 30–50 % coarse fraction is inverted (`IF(gravel>15, 4, 5)` and `IF(sand>15, 6, 7)`), so e.g. 5 % gravel gives "Sandy fat clay with gravel". `classify_ASTM2487` gives the correct names and needs no table.
* `slope_friction_circle`: translated line by line but **not yet checked against a worked example**. Please compare it with one of your workbooks.

## How it was checked

`tests/test.js` compares 45 results with textbook values (Prandtl/Terzaghi/Hansen/Vesić bearing factors, Coulomb and
Rankine coefficients, critical Coulomb wedge at 45° + φ/2, Terzaghi consolidation U(Tv), Boussinesq, Newmark and strip-load
stresses, Mohr circle, interpolation in increasing and decreasing order, shear/moment integration, ASTM D2487 groups),
and runs all 145 functions with realistic inputs to check that each returns a proper result.
The add-in has not yet been compared cell by cell with the LAMBDA library inside Excel.
