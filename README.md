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
- Formula evaluation engine with 20+ spreadsheet functions
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
| Math | `SUM`, `AVERAGE`, `MIN`, `MAX`, `ABS`, `ROUND`, `FLOOR`, `CEIL`, `INT`, `MOD`, `POWER`, `SQRT` |
| Logic | `IF`, `AND`, `OR`, `NOT`, `IFERROR` |
| Text | `CONCATENATE`, `LEN`, `TRIM`, `UPPER`, `LOWER`, `LEFT`, `RIGHT`, `MID`, `FIND`, `SEARCH`, `SUBSTITUTE`, `REPLACE`, `TEXT`, `VALUE` |
| Stats | `COUNT`, `COUNTA`, `COUNTIF`, `COUNTIFS` |
| Conditional | `SUMIF`, `SUMIFS`, `SUMPRODUCT` |
| Lookup | `VLOOKUP`, `HLOOKUP`, `INDEX`, `MATCH` |
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
		Family: "Arial",
		Size:   "12pt",
		Bold:   "bold",
		Color:  "#FF0000",
	},
	Fill: &ods.Fill{
		Color: "#FFFF00",
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

## Contributing

We welcome contributions! Please read our [Contributing Guide](CONTRIBUTING.md) and [Code of Conduct](CODE_OF_CONDUCT.md) before submitting a pull request.

## Security

To report a security vulnerability, please see our [Security Policy](SECURITY.md).

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
