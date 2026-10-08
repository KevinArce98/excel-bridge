# [1.5.0](https://github.com/KevinArce98/excel-bridge/compare/v1.4.0...v1.5.0) (2026-10-08)


### Bug Fixes

* **ci:** keep the release gate working with npm 12 ([9b257c7](https://github.com/KevinArce98/excel-bridge/commit/9b257c770d5686057483780c3b2fa5c8c8a7905f))
* **package:** resolve ESM types correctly and check the package in CI ([d789fef](https://github.com/KevinArce98/excel-bridge/commit/d789fefe9149ff853524c90b91f5f457340dd36f))
* **reader:** bound grid references, column ranges and zip inflation ([7f7d866](https://github.com/KevinArce98/excel-bridge/commit/7f7d8668150eb5a978759853275955c28488939f))
* **workbook:** keep rows, validations, number formats and sheet visibility on save ([2675051](https://github.com/KevinArce98/excel-bridge/commit/26750510a9577e8e15ec21a6a792140f0b9c216c))
* **writer:** reject input that produces files other tools cannot open ([69e7ed5](https://github.com/KevinArce98/excel-bridge/commit/69e7ed57e20c1dd1b86127632caf1b1446760cef))


### Features

* **workbook:** rename sheets and read error cells, and fix review findings ([8a5941d](https://github.com/KevinArce98/excel-bridge/commit/8a5941d08cb1edd756494831a4b3444704c7349a))


### Performance Improvements

* **writer:** visit only the populated rows of sparse sheets ([ec5e60e](https://github.com/KevinArce98/excel-bridge/commit/ec5e60e1c09b07e791f0f17c9af73b654e7c3a31))

# [1.4.0](https://github.com/KevinArce98/excel-bridge/compare/v1.3.0...v1.4.0) (2026-09-25)


### Bug Fixes

* **reader:** keep number-like cell text exact ([3b11239](https://github.com/KevinArce98/excel-bridge/commit/3b11239ed1dd33f6a1f3747261eabf727a99561e))
* **writer:** escape user-supplied attribute values ([f3e843b](https://github.com/KevinArce98/excel-bridge/commit/f3e843b580eb0273beb41523b3ef5edbe0e12edb))


### Features

* add autoFilter and hyperlinks ([a1b3750](https://github.com/KevinArce98/excel-bridge/commit/a1b3750e7e125370fe84a19d362bbd45648c3b13))

# [1.3.0](https://github.com/KevinArce98/excel-bridge/compare/v1.2.0...v1.3.0) (2026-08-29)


### Bug Fixes

* **writer:** emit worksheet and workbook elements in schema order ([b9743af](https://github.com/KevinArce98/excel-bridge/commit/b9743afeddbe51fa0cfa2d1e23ecd5a49ce3544c))


### Features

* **reader:** parse conditional formatting rules ([3648a42](https://github.com/KevinArce98/excel-bridge/commit/3648a421ea07d37c751b9e50b9f4eff8d236761e))
* **writer:** support whole, decimal, textLength and date validations ([4039317](https://github.com/KevinArce98/excel-bridge/commit/40393171c1ce916a09cc6fba6237e13929d564ba))

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
