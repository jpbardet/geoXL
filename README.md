# geoXL — student distribution

- [geoXL: install / update a workbook and function reference](https://jpbardet.github.io/geoXL/)

Use the updater to import the published LAMBDA library into a new copy of your
workbook. Save and close the original first. No add-in or macro installation is
required to update. Files are processed in your browser and are not uploaded.
Open the updated copy in Excel to calculate and check results.

The help describes arguments, returned values and existing usage notes. Theory,
references and validated worked examples will be expanded in later releases.
The textbook explains the teaching workbooks.

The updater reads `docs/library/current.json` and verifies the version and
SHA256 of its published library snapshot. It does not fetch the development
Gist. Current published version: 0.24 (2026-10-28); earlier snapshots are kept under `docs/library/releases/`, and a frozen copy of each release page under `docs/releases/`. Native Excel
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
