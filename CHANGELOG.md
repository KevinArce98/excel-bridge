# [1.2.0](https://github.com/KevinArce98/excel-bridge/compare/v1.1.0...v1.2.0) (2026-08-29)


### Features

* **reader:** read back cell styles, layout and document metadata ([fab3d47](https://github.com/KevinArce98/excel-bridge/commit/fab3d478fbe3d51dc6095570b6ae41864fb2214b))
* **types:** add conditional formatting types ([fea3a9b](https://github.com/KevinArce98/excel-bridge/commit/fea3a9bf4588a5052557581d42167159b1d08785))
* **workbook:** add high-level Workbook API ([3b37058](https://github.com/KevinArce98/excel-bridge/commit/3b3705816e1921845477ab97c34af197d28c5e89)), closes [hi#level](https://github.com/hi/issues/level)
* **writer:** add streaming workbook writer ([d0787d9](https://github.com/KevinArce98/excel-bridge/commit/d0787d9ff2a2a138a8b7fd5bba45deb96720dc19))
* **writer:** conditional formats, explicit column widths and docProps ([3162ac7](https://github.com/KevinArce98/excel-bridge/commit/3162ac7df974aad39fdce15a1388389e2d575e38))

# Changelog

All notable changes to this project are documented here. The format is based on
[Keep a Changelog](https://keepachangelog.com/) and this project adheres to
[Semantic Versioning](https://semver.org/).

## [1.1.0] - 2026-06-24

### Added
- **Shared strings (opt-in)** — `new ExcelWriter({ sharedStrings: true })` writes a
  shared-strings table instead of inline strings, reducing size for repeated text.
- **Extended cell styles** — `italic`, `underline`, `fontSize`, `fontName`, `align`,
  `verticalAlign`, `wrapText`, and custom `numberFormat` on `CellStyle`.
- **`CellValue` type** — the public writer API is now typed as `CellValue[][]`
  (`string | number | boolean | Date | null | undefined`) instead of `any[][]`.
- **Formula reading** — `ParsedCell.formula` exposes a cell's formula expression.
- **Date reading** — date-formatted cells are returned as `Date` objects with
  `type: 'date'`.

### Fixed
- **External file reading** — worksheets are now resolved through their relationship
  ids (`workbook.xml.rels`) instead of assuming `sheet{sheetId}.xml`, so files
  produced by Excel and other libraries load correctly.
- **Sparse rows** — cells are placed by their real column index, keeping columns
  aligned when a row omits empty cells.
- **Date timezone drift** — serial conversion now uses UTC calendar math, so dates
  round-trip to exact midnight instead of drifting by historical timezone offsets.
- **Formula cells** — no longer emit a misleading cached `<v>0</v>`; `fullCalcOnLoad`
  makes Excel recalculate on open.
- **Invalid XML characters** — control characters are stripped to prevent corrupt files.
- **`validateExcelStructure`** — accepts any worksheet part name, not only `sheet1.xml`.

### Changed
- Consolidated duplicate `CellStyle` / `CellValidation` definitions into
  `src/core/types.ts`.
- Added a `prepublishOnly` script (runs tests and build) and a `test:run` script.
