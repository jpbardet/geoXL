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
Gist. Current published version: 0.14 (2026-10-18); earlier snapshots are kept under `docs/library/releases/`. Native Excel
numerical and iterative-tool release validation remains pending.

Development tools, candidate testing and future release approval are maintained
separately. Historical JavaScript add-in files may remain in this repository for
later archiving; they are not required by the LAMBDA student updater.
