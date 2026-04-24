# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added
- Package-level godoc and per-symbol doc comments for every exported type,
  function, method, variable and constant.
- `examples/` folder with 7 runnable sample programs (quickstart, formulas,
  styling, autofilter, streaming, roundtrip, images).
- Typed cell helpers: `SetCellCurrency`, `SetCellPercentage`,
  `SetCellDateTime`, `SetCellDuration`.
- `.golangci.yml` and a golangci-lint CI job. CI matrix now covers Linux,
  macOS and Windows.

### Changed
- CI workflow: `go vet` step replaced by a golangci-lint-action run.

### Fixed
- `formula_eval.go`: misspelled `criterias` renamed, redundant nil-check and
  `WriteString(Sprintf)` pattern removed.

## [0.2.0] - 2026-04-24

### Added
- 14 formula functions: `MEDIAN`, `STDEV`, `STDEVP`, `VAR`, `VARP`, `RANK`,
  `LARGE`, `SMALL`, `PERCENTILE`, `RAND`, `RANDBETWEEN`, `IFS`, `SWITCH`,
  `XLOOKUP`.
- Advanced styling: text rotation, indent, super/subscript vertical
  alignment, strikethrough color, row auto-fit.
- Embedded images support (`AddImage`, `AddImageFromBytes`, `GetImages`,
  `RemoveImages`) with PNG/JPEG/GIF/BMP detection and SHA1-based
  deduplication.

## [0.1.2] - 2026-03-14

### Added
- Row/column visibility and auto-fit (`SetRowVisible`, `SetColVisible`,
  `SetRowAutoFit`, `SetColAutoFit`).
- Sheet protection (`SetSheetProtection`).
- Conditional formatting (`SetConditionalFormat` / `GetConditionalFormats`).
- Sort keys and filter criteria on auto-filters.

## [0.1.1] - 2026-03-13

### Added
- Initial formula engine with basic math, logic, text, stats and lookup
  functions.
- Hyperlinks on cells.
- CSV export (`ExportCSV`, `ExportCSVFile`).
- Per-cell number formats (`SetCellNumberFormat`, `GetCellFormattedValue`).
- Cross-sheet references in formulas.

## [0.1.0] - 2026-03-13

Initial release.

### Added
- Read/write of `.ods` files (zero external dependencies).
- Typed cell accessors (`SetCellStr`, `SetCellFloat`, `SetCellInt`,
  `SetCellBool`, `SetCellDate`).
- Sheet management, cell merging, cell styling, document properties.
- Data validation, named ranges, freeze panes, print ranges, page setup.
- Row iterator for streaming large sheets.

[Unreleased]: https://github.com/Mukbeast4/go-ods/compare/v0.2.0...HEAD
[0.2.0]: https://github.com/Mukbeast4/go-ods/compare/v0.1.2...v0.2.0
[0.1.2]: https://github.com/Mukbeast4/go-ods/compare/v0.1.1...v0.1.2
[0.1.1]: https://github.com/Mukbeast4/go-ods/compare/v0.1.0...v0.1.1
[0.1.0]: https://github.com/Mukbeast4/go-ods/releases/tag/v0.1.0
