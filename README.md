# geoXL — student distribution

- [geoXL: install / update a workbook and function reference](https://jpbardet.github.io/geoXL/) (geotechnical engineering course)
- [geolabXL: the laboratory library](https://jpbardet.github.io/geolabXL/) (soil mechanics laboratory)

Use the updater to import the published LAMBDA library into a new copy of your
workbook. Save and close the original first. No add-in or macro installation is
required to update. Files are processed in your browser and are not uploaded.
Open the updated copy in Excel to calculate and check results.

The help describes arguments, returned values and existing usage notes. Theory,
references and validated worked examples will be expanded in later releases.
The textbook explains the teaching workbooks.

The updater reads `docs/library/current.json` and verifies the version and
SHA256 of its published library snapshot. It does not fetch the development
Gist. Current published version: 0.25 (2026-11-05); earlier snapshots are kept under `docs/library/releases/`, and a frozen copy of each release page under `docs/releases/`. Native Excel
numerical and iterative-tool release validation remains pending.

Development tools, candidate testing and future release approval are maintained
separately. Historical JavaScript add-in files may remain in this repository for
later archiving; they are not required by the LAMBDA student updater.


Version 0.16 promotes the tested lateral-pile V2 integration under the original
function names. The updater migrates temporary _v2 names and preserves earlier
library snapshots. Independent analytical and browser checks passed; native
Excel numerical and iterative-tool validation remains pending.

Version 0.17 adds nine seepage_ functions for head interpolation, gradients, stream functions, contours/flow lines and segment profiles/statistics. Original formulas are preserved. The reference documents rectangular-grid, bottom-boundary and contour-plotting assumptions.

Version 0.22 revises bearing capacity: `bearing_capacity` takes `author` first and handles vertical, centred or eccentric loads; `bearing_capacity_inclined` handles inclined loads (Meyerhof, Hansen, Vesic, EC7 2004/2025).

Version 0.23 removes `bearing_inclinationFactors`; `bearing_capacity_inclined(..., out = 2)` returns the inclination factors. Workbooks that still call the old function keep their existing copy.

Version 0.24 adds nine mse_ functions for mechanically stabilized earth walls: earth pressure coefficient Kr (Rankine, AASHTO/FHWA simplified, coherent gravity), active zone (Rankine or bilinear), pullout factor F*, internal and external stability tables for ASD or LRFD, chart data, and a result reader. No existing function changed.

Version 0.25 adds 13 soil-laboratory functions and the separate laboratory page geolabXL 0.1. No existing function changed.

## Maintaining the two pages

There is one library and two separate student sites: geoXL (this repository, `docs/`, jpbardet.github.io/geoXL/) and geolabXL (repository `jpbardet/geolabXL`, `docs/`, jpbardet.github.io/geolabXL/). Both are built from this repository; `tools/build.py` writes the geolabXL site into a sibling checkout `../geolabXL` (or `GEOLABXL_DOCS`). Commit and push that repository separately.

| File | What it is |
|---|---|
| `library/geoXL.txt` | The master library. Every function is written here once. |
| `library/geolab.txt` | The names installed by the laboratory page. Functions they call are added automatically. |
| `documentation/function-help.txt` | Argument help, returns and method notes for both pages. |
| `documentation/geoxl-guide.html`, `geoxl-install-notes.html` | geoXL student guide and install-tab notices. |
| `documentation/geolab-guide.html` | geolabXL student guide. |
| `tools/build.py` | Builds both libraries, both pages and the frozen release copies (geoXL in `docs/`, geolabXL in `../geolabXL/docs/`). |
| `tools/updater-core.js`, `tools/page.js`, `tools/page.css` | The in-browser updater and page code shared by both pages. |
| `tests/updater/` | Runs the real updater in Chromium on folders of workbooks and checks the copies. |

To release:

1. Edit `library/geoXL.txt`: raise the version in its first line and in `version()`, and add a dated history line.
   If a function on the laboratory list (or one it calls) changed, also raise `version:` in `library/geolab.txt`.
2. Update the help text and the release notes in the guide(s).
3. Run `python3 tools/build.py`. It refuses to overwrite a published library snapshot.
4. Test: `TESTBED_DIRS='course=/path/to/course/workbooks' node tests/updater/run.js` then
   `python3 tests/updater/check.py` (see `tests/updater/run.js` for the laboratory page settings).
5. Commit and push `main`; GitHub Pages publishes `docs/`.

Each page stamps the workbooks it updates (`_GeoEXCEL_UpdateInfo` or `_GeoLabXL_UpdateInfo`). The geoXL page
refuses workbooks stamped only by geolabXL, so a laboratory workbook never receives the full library. The
laboratory page updates a geoXL workbook's laboratory functions and leaves the rest unchanged.
Example and laboratory workbooks are not kept in this repository; they are distributed through Canvas.
