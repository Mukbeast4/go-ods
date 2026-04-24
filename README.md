# go-ods

[![CI](https://github.com/Mukbeast4/go-ods/actions/workflows/ci.yml/badge.svg)](https://github.com/Mukbeast4/go-ods/actions/workflows/ci.yml)
[![Go Reference](https://pkg.go.dev/badge/github.com/mukbeast4/go-ods.svg)](https://pkg.go.dev/github.com/mukbeast4/go-ods)
[![Go Report Card](https://goreportcard.com/badge/github.com/mukbeast4/go-ods)](https://goreportcard.com/report/github.com/mukbeast4/go-ods)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Release](https://img.shields.io/github/v/release/Mukbeast4/go-ods)](https://github.com/Mukbeast4/go-ods/releases)

Pure Go library for reading, writing, and evaluating ODS (OpenDocument Spreadsheet) files. Zero external dependencies, no LibreOffice required.

## Features

- Read and write `.ods` files (OpenDocument Spreadsheet)
- Cell value management with typed getters/setters (string, int, float, bool, date)
- Formula evaluation engine with 45+ spreadsheet functions
- Formula recalculation with dependency graph and circular reference detection
- Auto-recalc mode: formulas update automatically when cell values change
- Sheet management (create, copy, rename, delete)
- Row and column operations (insert, remove, resize, hide/show)
- Column auto-fit (optimal width)
- Cell merging and unmerging
- Cell styling (font, fill, border, alignment)
- Sheet and cell protection
- Conditional formatting (calcext namespace)
- Auto-filter with filter criteria and sort keys
- Embedded images (PNG, JPEG, GIF, BMP) anchored to cells
- Streaming row iterator for large files
- Document properties (title, creator, description)

## Requirements

- Go 1.23+

## Installation

```bash
go get github.com/mukbeast4/go-ods
```

## Quick Start

```go
package main

import (
	"fmt"
	"log"

	ods "github.com/mukbeast4/go-ods"
)

func main() {
	f := ods.NewFile()

	f.SetCellValue("Sheet1", "A1", "Product")
	f.SetCellValue("Sheet1", "B1", "Price")
	f.SetCellFloat("Sheet1", "A2", 10)
	f.SetCellFloat("Sheet1", "A3", 20)
	f.SetCellFloat("Sheet1", "A4", 30)

	f.SetCellFormula("Sheet1", "A5", "SUM([.A2:.A4])")
	f.RecalcSheet("Sheet1")

	val, _ := f.GetCellFloat("Sheet1", "A5")
	fmt.Println("Sum:", val) // Sum: 60

	if err := f.SaveAs("output.ods"); err != nil {
		log.Fatal(err)
	}
}
```

## Formula Evaluation

The built-in formula engine supports evaluation without LibreOffice:

```go
f := ods.NewFile()

f.SetCellFloat("Sheet1", "A1", 100)
f.SetCellFormula("Sheet1", "B1", "[.A1]*1.2")
f.SetCellFormula("Sheet1", "C1", "IF([.B1]>100;\"over\";\"under\")")

f.RecalcAll()

b1, _ := f.GetCellFloat("Sheet1", "B1")   // 120
c1, _ := f.GetCellValue("Sheet1", "C1")   // "over"
```

### Auto-Recalc

Enable automatic formula recalculation when cell values change:

```go
f.SetAutoRecalc(true)

f.SetCellFloat("Sheet1", "A1", 50)
// All dependent formulas are recalculated immediately
```

### Supported Functions

| Category | Functions |
|----------|-----------|
| Math | `SUM`, `AVERAGE`, `MIN`, `MAX`, `ABS`, `ROUND`, `FLOOR`, `CEIL`, `INT`, `MOD`, `POWER`, `SQRT`, `RAND`, `RANDBETWEEN` |
| Logic | `IF`, `IFS`, `SWITCH`, `AND`, `OR`, `NOT`, `IFERROR` |
| Text | `CONCATENATE`, `LEN`, `TRIM`, `UPPER`, `LOWER`, `LEFT`, `RIGHT`, `MID`, `FIND`, `SEARCH`, `SUBSTITUTE`, `REPLACE`, `TEXT`, `VALUE` |
| Stats | `COUNT`, `COUNTA`, `COUNTIF`, `COUNTIFS`, `MEDIAN`, `STDEV`, `STDEVP`, `VAR`, `VARP`, `RANK`, `LARGE`, `SMALL`, `PERCENTILE` |
| Conditional | `SUMIF`, `SUMIFS`, `SUMPRODUCT` |
| Lookup | `VLOOKUP`, `HLOOKUP`, `XLOOKUP`, `INDEX`, `MATCH` |
| Date/Time | `DATE`, `TODAY`, `NOW`, `YEAR`, `MONTH`, `DAY` |

## Reading ODS Files

```go
f, err := ods.OpenFile("spreadsheet.ods")
if err != nil {
	log.Fatal(err)
}
defer f.Close()

val, _ := f.GetCellValue("Sheet1", "A1")
rows, _ := f.GetRows("Sheet1")

for _, row := range rows {
	fmt.Println(row)
}
```

## Sheet Operations

```go
f.NewSheet("Data")
f.SetSheetName("Sheet1", "Summary")
f.CopySheet("Data", "DataBackup")
f.DeleteSheet("DataBackup")

sheets := f.GetSheetList() // ["Summary", "Data"]
```

## Cell Styling

```go
style := &ods.Style{
	Font: &ods.Font{
		Family:             "Arial",
		Size:               "12pt",
		Bold:               "bold",
		Color:              "#FF0000",
		Strikethrough:      true,
		StrikethroughColor: "#000000",
		VerticalAlign:      "super", // or "sub"
	},
	Fill: &ods.Fill{Color: "#FFFF00"},
	Alignment: &ods.Alignment{
		Horizontal: "center",
		Rotation:   45,  // degrees (0-360, negatives are normalized)
		Indent:     2,   // 0.25cm increments
	},
}

styleID, _ := f.NewStyle(style)
f.SetCellStyle("Sheet1", "A1", "A1", styleID)
```

## Row and Column Operations

```go
f.InsertRows("Sheet1", 3, 2)       // Insert 2 rows at row 3
f.RemoveRow("Sheet1", 5)           // Remove row 5
f.SetColWidth("Sheet1", "B", 5.0)  // Set column B width

f.SetRowVisible("Sheet1", 2, false)  // Hide row 2
f.SetColVisible("Sheet1", "C", false) // Hide column C
f.SetColAutoFit("Sheet1", "A", true)  // Auto-fit column A width
f.SetRowAutoFit("Sheet1", 1, true)    // Auto-fit row 1 height
```

## Sheet Protection

```go
f.SetSheetProtection("Sheet1", true)

protected, _ := f.IsSheetProtected("Sheet1") // true

locked := true
styleID, _ := f.NewStyle(&ods.Style{
	Protected: &locked,
})
f.SetCellStyle("Sheet1", "A1", "A1", styleID)
```

## Conditional Formatting

```go
rules := []ods.ConditionalRule{
	{Value: "of:cell-content()>90", StyleName: "good", BaseCellAddress: "Sheet1.A1"},
	{Value: "of:cell-content()<50", StyleName: "bad", BaseCellAddress: "Sheet1.A1"},
}
f.SetConditionalFormat("Sheet1", "A1:A100", rules)

formats, _ := f.GetConditionalFormats("Sheet1")
f.RemoveConditionalFormat("Sheet1", "A1:A100")
```

## Auto-Filter, Filter Criteria and Sort

```go
f.SetAutoFilter("Sheet1", "A1", "D100")

f.SetFilterCriteria("Sheet1", []ods.FilterCriteria{
	{Column: 0, Values: []string{"Alice"}},
})

f.SetSort("Sheet1", []ods.SortKey{
	{Column: 1, Descending: true},
})

criteria, _ := f.GetFilterCriteria("Sheet1")
sortKeys, _ := f.GetSort("Sheet1")

f.ClearFilterCriteria("Sheet1")
f.RemoveSort("Sheet1")
```

## Embedded Images

Anchor images (PNG, JPEG, GIF, BMP) to any cell. Format is auto-detected, dimensions are in centimeters, identical binaries are deduplicated.

```go
f.AddImage("Sheet1", "B2", "logo.png", &ods.ImageOptions{Width: 4, Height: 3})

data, _ := os.ReadFile("chart.jpg")
f.AddImageFromBytes("Sheet1", "D5", data, &ods.ImageOptions{
	Width: 6, Height: 4, OffsetX: 0.5, OffsetY: 0.25,
})

images, _ := f.GetImages("Sheet1")
for _, img := range images {
	fmt.Printf("%s: %s %vx%vcm\n", img.CellRef, img.Format, img.Width, img.Height)
}

f.RemoveImages("Sheet1", "B2")
```

## Examples

A collection of runnable samples lives under [`examples/`](examples). Each sub-directory is a standalone `main` package:

```bash
go run ./examples/quickstart   # cells, formula, save
go run ./examples/formulas     # aggregation functions
go run ./examples/styling      # fonts, fills, borders, rotation
go run ./examples/autofilter   # filter + sort
go run ./examples/streaming    # 10k rows via AppendRows + RowIterator
go run ./examples/roundtrip    # open -> mutate -> save
go run ./examples/images       # embed a PNG
```

See [examples/README.md](examples/README.md) for a breakdown of each sample.

## Contributing

We welcome contributions! Please read our [Contributing Guide](CONTRIBUTING.md) and [Code of Conduct](CODE_OF_CONDUCT.md) before submitting a pull request.

## Security

To report a security vulnerability, please see our [Security Policy](SECURITY.md).

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
