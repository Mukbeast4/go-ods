//go:build ignore

package main

import (
	"fmt"
	"log"
	"time"

	ods "github.com/mukbeast4/go-ods"
)

func main() {
	generators := []struct {
		name string
		fn   func() error
	}{
		{"golden_types.ods", genGoldenTypes},
		{"golden_styles.ods", genGoldenStyles},
		{"golden_formulas.ods", genGoldenFormulas},
		{"golden_merge.ods", genGoldenMerge},
		{"golden_features.ods", genGoldenFeatures},
		{"golden_large.ods", genGoldenLarge},
		{"golden_multisheet.ods", genGoldenMultisheet},
	}

	for _, g := range generators {
		if err := g.fn(); err != nil {
			log.Fatalf("generate %s: %v", g.name, err)
		}
		log.Printf("generated testdata/%s", g.name)
	}
}

func genGoldenTypes() error {
	f := ods.NewFile()

	f.SetCellStr("Sheet1", "A1", "Type")
	f.SetCellStr("Sheet1", "B1", "Value")

	f.SetCellStr("Sheet1", "A2", "string")
	f.SetCellStr("Sheet1", "B2", "hello world")

	f.SetCellStr("Sheet1", "A3", "empty_string")
	f.SetCellStr("Sheet1", "B3", "")

	f.SetCellStr("Sheet1", "A4", "float")
	f.SetCellFloat("Sheet1", "B4", 3.14159)

	f.SetCellStr("Sheet1", "A5", "int")
	f.SetCellInt("Sheet1", "B5", 42)

	f.SetCellStr("Sheet1", "A6", "negative_float")
	f.SetCellFloat("Sheet1", "B6", -99.5)

	f.SetCellStr("Sheet1", "A7", "zero")
	f.SetCellFloat("Sheet1", "B7", 0)

	f.SetCellStr("Sheet1", "A8", "bool_true")
	f.SetCellBool("Sheet1", "B8", true)

	f.SetCellStr("Sheet1", "A9", "bool_false")
	f.SetCellBool("Sheet1", "B9", false)

	f.SetCellStr("Sheet1", "A10", "date")
	f.SetCellDate("Sheet1", "B10", time.Date(2024, 6, 15, 0, 0, 0, 0, time.UTC))

	f.SetCellStr("Sheet1", "A11", "date_epoch")
	f.SetCellDate("Sheet1", "B11", time.Date(1970, 1, 1, 0, 0, 0, 0, time.UTC))

	f.SetCellStr("Sheet1", "A12", "large_int")
	f.SetCellInt("Sheet1", "B12", 999999999)

	f.SetCellStr("Sheet1", "A13", "small_float")
	f.SetCellFloat("Sheet1", "B13", 0.000001)

	f.SetCellStr("Sheet1", "A14", "large_float")
	f.SetCellFloat("Sheet1", "B14", 1234567.890123)

	f.SetCellStr("Sheet1", "A15", "string_numeric")
	f.SetCellStr("Sheet1", "B15", "12345")

	f.SetCellStr("Sheet1", "A16", "string_special")
	f.SetCellStr("Sheet1", "B16", "line1\nline2")

	f.SetCellStr("Sheet1", "A17", "string_unicode")
	f.SetCellStr("Sheet1", "B17", "cafe\u0301")

	f.SetCellStr("Sheet1", "A18", "negative_int")
	f.SetCellInt("Sheet1", "B18", -100)

	f.SetCellStr("Sheet1", "A19", "max_float")
	f.SetCellFloat("Sheet1", "B19", 1.7976931348623157e+100)

	f.SetCellStr("Sheet1", "A20", "bool_string")
	f.SetCellStr("Sheet1", "B20", "TRUE")

	f.NewSheet("Sheet2")
	f.SetCellStr("Sheet2", "A1", "Sheet2 String")
	f.SetCellFloat("Sheet2", "A2", 100)
	f.SetCellBool("Sheet2", "A3", true)
	f.SetCellDate("Sheet2", "A4", time.Date(2025, 12, 31, 0, 0, 0, 0, time.UTC))

	return f.SaveAs("testdata/golden_types.ods")
}

func genGoldenStyles() error {
	f := ods.NewFile()

	boldID, _ := f.NewStyle(&ods.Style{
		Font: &ods.Font{Bold: "bold"},
	})
	italicID, _ := f.NewStyle(&ods.Style{
		Font: &ods.Font{Italic: "italic"},
	})
	underlineID, _ := f.NewStyle(&ods.Style{
		Font: &ods.Font{Underline: true},
	})
	boldItalicID, _ := f.NewStyle(&ods.Style{
		Font: &ods.Font{Bold: "bold", Italic: "italic"},
	})
	fillRedID, _ := f.NewStyle(&ods.Style{
		Fill: &ods.Fill{Color: "#FF0000"},
	})
	fillGreenID, _ := f.NewStyle(&ods.Style{
		Fill: &ods.Fill{Color: "#00FF00"},
	})
	borderID, _ := f.NewStyle(&ods.Style{
		Border: &ods.Border{Style: "solid", Width: "0.06pt", Color: "#000000"},
	})
	alignCenterID, _ := f.NewStyle(&ods.Style{
		Alignment: &ods.Alignment{Horizontal: "center"},
	})
	alignRightID, _ := f.NewStyle(&ods.Style{
		Alignment: &ods.Alignment{Horizontal: "end"},
	})
	wrapID, _ := f.NewStyle(&ods.Style{
		Alignment: &ods.Alignment{WrapText: true},
	})
	fontSizeID, _ := f.NewStyle(&ods.Style{
		Font: &ods.Font{Size: "14pt"},
	})
	fontColorID, _ := f.NewStyle(&ods.Style{
		Font: &ods.Font{Color: "#0000FF"},
	})
	strikeID, _ := f.NewStyle(&ods.Style{
		Font: &ods.Font{Strikethrough: true},
	})
	comboID, _ := f.NewStyle(&ods.Style{
		Font:      &ods.Font{Bold: "bold", Color: "#FFFFFF"},
		Fill:      &ods.Fill{Color: "#333333"},
		Border:    &ods.Border{Style: "solid", Width: "1pt", Color: "#FF0000"},
		Alignment: &ods.Alignment{Horizontal: "center", Vertical: "middle"},
	})
	verticalID, _ := f.NewStyle(&ods.Style{
		Alignment: &ods.Alignment{Vertical: "bottom"},
	})

	styles := []struct {
		label   string
		styleID int
	}{
		{"Bold", boldID},
		{"Italic", italicID},
		{"Underline", underlineID},
		{"Bold+Italic", boldItalicID},
		{"Fill Red", fillRedID},
		{"Fill Green", fillGreenID},
		{"Border Solid", borderID},
		{"Align Center", alignCenterID},
		{"Align Right", alignRightID},
		{"Wrap Text", wrapID},
		{"Font Size 14", fontSizeID},
		{"Font Blue", fontColorID},
		{"Strikethrough", strikeID},
		{"Combo Style", comboID},
		{"Vertical Bottom", verticalID},
	}

	f.SetCellStr("Sheet1", "A1", "Style")
	f.SetCellStr("Sheet1", "B1", "Sample")

	for i, s := range styles {
		row := i + 2
		cell := fmt.Sprintf("A%d", row)
		valCell := fmt.Sprintf("B%d", row)
		f.SetCellStr("Sheet1", cell, s.label)
		f.SetCellStr("Sheet1", valCell, s.label+" value")
		f.SetCellStyle("Sheet1", valCell, valCell, s.styleID)
	}

	return f.SaveAs("testdata/golden_styles.ods")
}

func genGoldenFormulas() error {
	f := ods.NewFile()

	f.SetCellFloat("Sheet1", "A1", 10)
	f.SetCellFloat("Sheet1", "A2", 20)
	f.SetCellFloat("Sheet1", "A3", 30)
	f.SetCellFloat("Sheet1", "A4", 40)
	f.SetCellFloat("Sheet1", "A5", 50)

	f.SetCellStr("Sheet1", "B1", "SUM")
	f.SetCellFormula("Sheet1", "C1", "SUM([.A1:.A5])")

	f.SetCellStr("Sheet1", "B2", "AVERAGE")
	f.SetCellFormula("Sheet1", "C2", "AVERAGE([.A1:.A5])")

	f.SetCellStr("Sheet1", "B3", "COUNT")
	f.SetCellFormula("Sheet1", "C3", "COUNT([.A1:.A5])")

	f.SetCellStr("Sheet1", "B4", "MIN")
	f.SetCellFormula("Sheet1", "C4", "MIN([.A1:.A5])")

	f.SetCellStr("Sheet1", "B5", "MAX")
	f.SetCellFormula("Sheet1", "C5", "MAX([.A1:.A5])")

	f.SetCellStr("Sheet1", "B6", "IF_true")
	f.SetCellFormula("Sheet1", "C6", "IF([.A1]>5;\"big\";\"small\")")

	f.SetCellStr("Sheet1", "B7", "IF_false")
	f.SetCellFormula("Sheet1", "C7", "IF([.A1]<5;\"big\";\"small\")")

	f.SetCellStr("Sheet1", "B8", "arithmetic")
	f.SetCellFormula("Sheet1", "C8", "[.A1]+[.A2]*[.A3]")

	f.SetCellStr("Sheet1", "B9", "ABS")
	f.SetCellFormula("Sheet1", "C9", "ABS(-42)")

	f.SetCellStr("Sheet1", "B10", "ROUND")
	f.SetCellFormula("Sheet1", "C10", "ROUND(3.14159;2)")

	f.SetCellStr("Sheet1", "B11", "CONCATENATE")
	f.SetCellFormula("Sheet1", "C11", "CONCATENATE(\"hello\";\" \";\"world\")")

	f.SetCellStr("Sheet1", "B12", "LEN")
	f.SetCellFormula("Sheet1", "C12", "LEN(\"hello\")")

	f.SetCellStr("Sheet1", "B13", "UPPER")
	f.SetCellFormula("Sheet1", "C13", "UPPER(\"hello\")")

	f.SetCellStr("Sheet1", "B14", "LOWER")
	f.SetCellFormula("Sheet1", "C14", "LOWER(\"HELLO\")")

	f.SetCellStr("Sheet1", "B15", "SUMIF")
	f.SetCellFormula("Sheet1", "C15", "SUMIF([.A1:.A5];\">20\")")

	f.NewSheet("Data")
	f.SetCellStr("Data", "A1", "Name")
	f.SetCellStr("Data", "B1", "Score")
	f.SetCellStr("Data", "A2", "Alice")
	f.SetCellFloat("Data", "B2", 85)
	f.SetCellStr("Data", "A3", "Bob")
	f.SetCellFloat("Data", "B3", 72)
	f.SetCellStr("Data", "A4", "Charlie")
	f.SetCellFloat("Data", "B4", 91)
	f.SetCellStr("Data", "A5", "Diana")
	f.SetCellFloat("Data", "B5", 68)

	f.NewSheet("CrossRef")
	f.SetCellStr("CrossRef", "A1", "DataSum")
	f.SetCellFormula("CrossRef", "B1", "SUM([Data.B2:.B5])")

	f.SetCellStr("CrossRef", "A2", "DataAvg")
	f.SetCellFormula("CrossRef", "B2", "AVERAGE([Data.B2:.B5])")

	f.SetCellStr("CrossRef", "A3", "Sheet1Val")
	f.SetCellFormula("CrossRef", "B3", "[Sheet1.A1]+[Sheet1.A2]")

	f.SetCellStr("CrossRef", "A4", "VLOOKUP")
	f.SetCellFormula("CrossRef", "B4", "VLOOKUP(\"Bob\";[Data.A2:.B5];2;0)")

	f.SetCellStr("CrossRef", "A5", "COUNTIF")
	f.SetCellFormula("CrossRef", "B5", "COUNTIF([Data.B2:.B5];\">75\")")

	f.RecalcAll()

	return f.SaveAs("testdata/golden_formulas.ods")
}

func genGoldenMerge() error {
	f := ods.NewFile()

	f.SetCellStr("Sheet1", "A1", "Merged 2x2")
	f.MergeCell("Sheet1", "A1", "B2")

	f.SetCellStr("Sheet1", "A4", "Merged 1x5")
	f.MergeCell("Sheet1", "A4", "E4")

	f.SetCellStr("Sheet1", "D1", "Merged 3x1")
	f.MergeCell("Sheet1", "D1", "D3")

	f.SetCellStr("Sheet1", "A6", "Not merged")
	f.SetCellStr("Sheet1", "B6", "Also not merged")

	return f.SaveAs("testdata/golden_merge.ods")
}

func genGoldenFeatures() error {
	f := ods.NewFile()

	f.SetCellStr("Sheet1", "A1", "With Comment")
	f.SetCellComment("Sheet1", "A1", &ods.Comment{
		Author: "TestUser",
		Date:   "2024-01-15T10:00:00",
		Text:   "This is a test comment",
	})

	f.SetCellStr("Sheet1", "A2", "Multi-line Comment")
	f.SetCellComment("Sheet1", "A2", &ods.Comment{
		Author: "TestUser",
		Date:   "2024-01-15T11:00:00",
		Text:   "Line 1\nLine 2\nLine 3",
	})

	f.SetCellStr("Sheet1", "A3", "Example Link")
	f.SetCellHyperlink("Sheet1", "A3", "https://example.com", "Example Link")

	f.SetCellStr("Sheet1", "A4", "Go Link")
	f.SetCellHyperlink("Sheet1", "A4", "https://go.dev", "Go Link")

	f.SetNamedRange("TestRange", "Sheet1", "A1", "B5")
	f.SetNamedRange("SingleCell", "Sheet1", "C1", "C1")

	f.SetCellInt("Sheet1", "B1", 50)
	f.SetCellInt("Sheet1", "B2", 75)
	f.SetCellInt("Sheet1", "B3", 25)
	f.SetCellInt("Sheet1", "B4", 90)
	f.SetCellInt("Sheet1", "B5", 10)

	f.SetDataValidation("Sheet1", "B1", "B5", &ods.DataValidation{
		Type:         "whole-number",
		Operator:     "between",
		Formula1:     "1",
		Formula2:     "100",
		AllowEmpty:   true,
		ErrorTitle:   "Invalid",
		ErrorMessage: "Enter 1-100",
		ErrorStyle:   "stop",
		InputTitle:   "Input",
		InputMessage: "Enter a number",
	})

	f.SetAutoFilter("Sheet1", "A1", "D5")

	f.SetFreezePane("Sheet1", "B2")

	return f.SaveAs("testdata/golden_features.ods")
}

func genGoldenLarge() error {
	f := ods.NewFile()

	for col := 1; col <= 26; col++ {
		colName, _ := ods.CoordinatesToCellName(col, 1)
		f.SetCellStr("Sheet1", colName, fmt.Sprintf("Col%c", 'A'+col-1))
	}

	for row := 2; row <= 1001; row++ {
		for col := 1; col <= 26; col++ {
			cellName, _ := ods.CoordinatesToCellName(col, row)
			if col%2 == 0 {
				f.SetCellFloat("Sheet1", cellName, float64((row-1)*100+col))
			} else {
				f.SetCellStr("Sheet1", cellName, fmt.Sprintf("R%dC%d", row-1, col))
			}
		}
	}

	return f.SaveAs("testdata/golden_large.ods")
}

func genGoldenMultisheet() error {
	f := ods.NewFile()

	f.SetCellStr("Sheet1", "A1", "Sheet1 Data")
	for i := 1; i <= 10; i++ {
		f.SetCellFloat("Sheet1", fmt.Sprintf("A%d", i+1), float64(i*10))
	}

	f.NewSheet("Sheet2")
	f.SetCellStr("Sheet2", "A1", "Sheet2 Data")
	for i := 1; i <= 10; i++ {
		f.SetCellFloat("Sheet2", fmt.Sprintf("A%d", i+1), float64(i*5))
	}

	f.NewSheet("Sheet3")
	f.SetCellStr("Sheet3", "A1", "Cross-Sheet Sums")
	f.SetCellFormula("Sheet3", "A2", "SUM([Sheet1.A2:.A11])")
	f.SetCellFormula("Sheet3", "A3", "SUM([Sheet2.A2:.A11])")
	f.SetCellFormula("Sheet3", "A4", "[Sheet1.A2]+[Sheet2.A2]")

	f.NewSheet("Sheet4")
	f.SetCellStr("Sheet4", "A1", "Lookups")
	f.SetCellStr("Sheet4", "B1", "Key")
	f.SetCellStr("Sheet4", "C1", "Value")
	keys := []string{"alpha", "beta", "gamma", "delta", "epsilon"}
	for i, k := range keys {
		f.SetCellStr("Sheet4", fmt.Sprintf("B%d", i+2), k)
		f.SetCellFloat("Sheet4", fmt.Sprintf("C%d", i+2), float64((i+1)*100))
	}

	f.NewSheet("Sheet5")
	f.SetCellStr("Sheet5", "A1", "Summary")
	f.SetCellFormula("Sheet5", "A2", "[Sheet3.A2]+[Sheet3.A3]")
	f.SetCellFormula("Sheet5", "A3", "AVERAGE([Sheet1.A2:.A11])")

	f.SetNamedRange("Sheet1Data", "Sheet1", "A2", "A11")
	f.SetNamedRange("Sheet2Data", "Sheet2", "A2", "A11")

	f.RecalcAll()

	return f.SaveAs("testdata/golden_multisheet.ods")
}
