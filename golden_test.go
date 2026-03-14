package goods

import (
	"math"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func loadGolden(t *testing.T, name string) *File {
	t.Helper()
	path := filepath.Join("testdata", name)
	f, err := OpenFile(path)
	if err != nil {
		t.Fatalf("load golden %s: %v", name, err)
	}
	return f
}

func saveAndReopen(t *testing.T, f *File) *File {
	t.Helper()
	dir, err := os.MkdirTemp("", "goods-golden-*")
	if err != nil {
		t.Fatalf("create temp dir: %v", err)
	}
	t.Cleanup(func() { os.RemoveAll(dir) })

	path := filepath.Join(dir, "roundtrip.ods")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("save: %v", err)
	}

	reopened, err := OpenFile(path)
	if err != nil {
		t.Fatalf("reopen: %v", err)
	}
	return reopened
}

func assertCellValue(t *testing.T, f *File, sheet, cell, expected string) {
	t.Helper()
	got, err := f.GetCellValue(sheet, cell)
	if err != nil {
		t.Errorf("GetCellValue(%s, %s): %v", sheet, cell, err)
		return
	}
	if got != expected {
		t.Errorf("GetCellValue(%s, %s) = %q, want %q", sheet, cell, got, expected)
	}
}

func assertCellFloat(t *testing.T, f *File, sheet, cell string, expected float64) {
	t.Helper()
	got, err := f.GetCellFloat(sheet, cell)
	if err != nil {
		t.Errorf("GetCellFloat(%s, %s): %v", sheet, cell, err)
		return
	}
	if math.Abs(got-expected) > 1e-9 {
		t.Errorf("GetCellFloat(%s, %s) = %v, want %v", sheet, cell, got, expected)
	}
}

func assertCellFormula(t *testing.T, f *File, sheet, cell, expected string) {
	t.Helper()
	got, err := f.GetCellFormula(sheet, cell)
	if err != nil {
		t.Errorf("GetCellFormula(%s, %s): %v", sheet, cell, err)
		return
	}
	if got != expected {
		t.Errorf("GetCellFormula(%s, %s) = %q, want %q", sheet, cell, got, expected)
	}
}

func assertCellType(t *testing.T, f *File, sheet, cell string, expected CellType) {
	t.Helper()
	got, err := f.GetCellType(sheet, cell)
	if err != nil {
		t.Errorf("GetCellType(%s, %s): %v", sheet, cell, err)
		return
	}
	if got != expected {
		t.Errorf("GetCellType(%s, %s) = %d, want %d", sheet, cell, got, expected)
	}
}

func verifyTypesFile(t *testing.T, f *File) {
	t.Helper()

	assertCellValue(t, f, "Sheet1", "B2", "hello world")
	assertCellType(t, f, "Sheet1", "B2", CellTypeString)

	assertCellFloat(t, f, "Sheet1", "B4", 3.14159)
	assertCellFloat(t, f, "Sheet1", "B5", 42)
	assertCellFloat(t, f, "Sheet1", "B6", -99.5)
	assertCellFloat(t, f, "Sheet1", "B7", 0)

	assertCellValue(t, f, "Sheet1", "B8", "TRUE")
	assertCellType(t, f, "Sheet1", "B8", CellTypeBool)
	assertCellValue(t, f, "Sheet1", "B9", "FALSE")
	assertCellType(t, f, "Sheet1", "B9", CellTypeBool)

	dateVal, err := f.GetCellValue("Sheet1", "B10")
	if err != nil {
		t.Errorf("GetCellValue(Sheet1, B10): %v", err)
	} else if !strings.Contains(dateVal, "2024-06-15") {
		t.Errorf("B10 date = %q, want contains '2024-06-15'", dateVal)
	}

	assertCellFloat(t, f, "Sheet1", "B12", 999999999)

	assertCellValue(t, f, "Sheet1", "B15", "12345")
	assertCellType(t, f, "Sheet1", "B15", CellTypeString)

	assertCellValue(t, f, "Sheet2", "A1", "Sheet2 String")
	assertCellFloat(t, f, "Sheet2", "A2", 100)
	assertCellType(t, f, "Sheet2", "A3", CellTypeBool)
}

func TestGoldenRoundTripTypes(t *testing.T) {
	f := loadGolden(t, "golden_types.ods")

	t.Run("before_roundtrip", func(t *testing.T) {
		verifyTypesFile(t, f)
	})

	f2 := saveAndReopen(t, f)

	t.Run("after_roundtrip", func(t *testing.T) {
		verifyTypesFile(t, f2)
	})
}

func TestGoldenRoundTripFormulas(t *testing.T) {
	f := loadGolden(t, "golden_formulas.ods")

	t.Run("formula_text_preserved", func(t *testing.T) {
		assertCellFormula(t, f, "Sheet1", "C1", "SUM([.A1:.A5])")
		assertCellFormula(t, f, "Sheet1", "C2", "AVERAGE([.A1:.A5])")

		c6Formula, err := f.GetCellFormula("Sheet1", "C6")
		if err != nil {
			t.Errorf("GetCellFormula(Sheet1, C6): %v", err)
		} else if !strings.Contains(c6Formula, "IF") {
			t.Errorf("C6 formula = %q, want contains 'IF'", c6Formula)
		}

		b1Formula, err := f.GetCellFormula("CrossRef", "B1")
		if err != nil {
			t.Errorf("GetCellFormula(CrossRef, B1): %v", err)
		} else if !strings.Contains(b1Formula, "SUM") {
			t.Errorf("CrossRef B1 formula = %q, want contains 'SUM'", b1Formula)
		}

		b3Formula, err := f.GetCellFormula("CrossRef", "B3")
		if err != nil {
			t.Errorf("GetCellFormula(CrossRef, B3): %v", err)
		} else if !strings.Contains(b3Formula, "Sheet1") {
			t.Errorf("CrossRef B3 formula = %q, want contains 'Sheet1'", b3Formula)
		}
	})

	t.Run("values_before_roundtrip", func(t *testing.T) {
		assertCellValue(t, f, "Sheet1", "C1", "150")
	})

	f2 := saveAndReopen(t, f)

	t.Run("formulas_after_roundtrip", func(t *testing.T) {
		assertCellFormula(t, f2, "Sheet1", "C1", "SUM([.A1:.A5])")
		assertCellFormula(t, f2, "Sheet1", "C2", "AVERAGE([.A1:.A5])")

		c6Formula, err := f2.GetCellFormula("Sheet1", "C6")
		if err != nil {
			t.Errorf("GetCellFormula(Sheet1, C6): %v", err)
		} else if !strings.Contains(c6Formula, "IF") {
			t.Errorf("C6 formula = %q, want contains 'IF'", c6Formula)
		}

		b1Formula, err := f2.GetCellFormula("CrossRef", "B1")
		if err != nil {
			t.Errorf("GetCellFormula(CrossRef, B1): %v", err)
		} else if !strings.Contains(b1Formula, "SUM") {
			t.Errorf("CrossRef B1 formula = %q, want contains 'SUM'", b1Formula)
		}

		b3Formula, err := f2.GetCellFormula("CrossRef", "B3")
		if err != nil {
			t.Errorf("GetCellFormula(CrossRef, B3): %v", err)
		} else if !strings.Contains(b3Formula, "Sheet1") {
			t.Errorf("CrossRef B3 formula = %q, want contains 'Sheet1'", b3Formula)
		}
	})

	t.Run("recalc_after_roundtrip", func(t *testing.T) {
		if err := f2.RecalcAll(); err != nil {
			t.Fatalf("RecalcAll: %v", err)
		}
		assertCellFloat(t, f2, "Sheet1", "C1", 150)
	})
}

func TestGoldenRoundTripMerge(t *testing.T) {
	f := loadGolden(t, "golden_merge.ods")

	mergesBefore, err := f.GetMergeCells("Sheet1")
	if err != nil {
		t.Fatalf("GetMergeCells before: %v", err)
	}

	f2 := saveAndReopen(t, f)

	mergesAfter, err := f2.GetMergeCells("Sheet1")
	if err != nil {
		t.Fatalf("GetMergeCells after: %v", err)
	}

	if len(mergesBefore) != len(mergesAfter) {
		t.Fatalf("merge count: before=%d after=%d", len(mergesBefore), len(mergesAfter))
	}

	beforeSet := make(map[[2]string]bool)
	for _, m := range mergesBefore {
		beforeSet[m] = true
	}
	for _, m := range mergesAfter {
		if !beforeSet[m] {
			t.Errorf("merge %v present after roundtrip but not before", m)
		}
	}
}

func TestGoldenRoundTripFeatures(t *testing.T) {
	f := loadGolden(t, "golden_features.ods")

	verifyFeatures := func(t *testing.T, f *File, label string) {
		t.Helper()

		t.Run(label+"/comments", func(t *testing.T) {
			c1, err := f.GetCellComment("Sheet1", "A1")
			if err != nil {
				t.Fatalf("GetCellComment A1: %v", err)
			}
			if c1 == nil {
				t.Fatal("A1 comment is nil")
			}
			if c1.Author != "TestUser" {
				t.Errorf("A1 comment author = %q, want 'TestUser'", c1.Author)
			}
			if c1.Text != "This is a test comment" {
				t.Errorf("A1 comment text = %q, want 'This is a test comment'", c1.Text)
			}

			c2, err := f.GetCellComment("Sheet1", "A2")
			if err != nil {
				t.Fatalf("GetCellComment A2: %v", err)
			}
			if c2 == nil {
				t.Fatal("A2 comment is nil")
			}
			if !strings.Contains(c2.Text, "Line 1") {
				t.Errorf("A2 comment text = %q, want contains 'Line 1'", c2.Text)
			}
		})

		t.Run(label+"/hyperlinks", func(t *testing.T) {
			url, display, err := f.GetCellHyperlink("Sheet1", "A3")
			if err != nil {
				t.Fatalf("GetCellHyperlink A3: %v", err)
			}
			if url != "https://example.com" {
				t.Errorf("A3 hyperlink url = %q, want 'https://example.com'", url)
			}
			if display != "Example Link" {
				t.Errorf("A3 hyperlink display = %q, want 'Example Link'", display)
			}

			url2, _, err := f.GetCellHyperlink("Sheet1", "A4")
			if err != nil {
				t.Fatalf("GetCellHyperlink A4: %v", err)
			}
			if url2 != "https://go.dev" {
				t.Errorf("A4 hyperlink url = %q, want 'https://go.dev'", url2)
			}
		})

		t.Run(label+"/named_ranges", func(t *testing.T) {
			ranges := f.GetNamedRanges()
			if len(ranges) < 2 {
				t.Fatalf("named ranges count = %d, want >= 2", len(ranges))
			}
			found := map[string]bool{}
			for _, nr := range ranges {
				found[nr.Name] = true
			}
			if !found["TestRange"] {
				t.Error("named range 'TestRange' not found")
			}
			if !found["SingleCell"] {
				t.Error("named range 'SingleCell' not found")
			}
		})

		t.Run(label+"/validations", func(t *testing.T) {
			dv, err := f.GetDataValidation("Sheet1", "B1")
			if err != nil {
				t.Fatalf("GetDataValidation B1: %v", err)
			}
			if dv == nil {
				t.Fatal("B1 validation is nil")
			}
			if dv.Type != "whole-number" {
				t.Errorf("B1 validation type = %q, want 'whole-number'", dv.Type)
			}
			if dv.Operator != "between" {
				t.Errorf("B1 validation operator = %q, want 'between'", dv.Operator)
			}
			if dv.Formula1 != "1" {
				t.Errorf("B1 validation formula1 = %q, want '1'", dv.Formula1)
			}
			if dv.Formula2 != "100" {
				t.Errorf("B1 validation formula2 = %q, want '100'", dv.Formula2)
			}
		})

		t.Run(label+"/autofilter", func(t *testing.T) {
			tl, br, err := f.GetAutoFilter("Sheet1")
			if err != nil {
				t.Fatalf("GetAutoFilter: %v", err)
			}
			if tl != "A1" {
				t.Errorf("autofilter top-left = %q, want 'A1'", tl)
			}
			if br != "D5" {
				t.Errorf("autofilter bottom-right = %q, want 'D5'", br)
			}
		})
	}

	verifyFeatures(t, f, "before_roundtrip")

	f2 := saveAndReopen(t, f)

	verifyFeatures(t, f2, "after_roundtrip")
}

func TestGoldenRoundTripStyles(t *testing.T) {
	f := loadGolden(t, "golden_styles.ods")

	assertCellValue(t, f, "Sheet1", "B2", "Bold value")
	assertCellValue(t, f, "Sheet1", "B3", "Italic value")
	assertCellValue(t, f, "Sheet1", "B15", "Combo Style value")

	f2 := saveAndReopen(t, f)

	assertCellValue(t, f2, "Sheet1", "B2", "Bold value")
	assertCellValue(t, f2, "Sheet1", "B3", "Italic value")
	assertCellValue(t, f2, "Sheet1", "B15", "Combo Style value")
}

func TestGoldenRoundTripMultisheet(t *testing.T) {
	f := loadGolden(t, "golden_multisheet.ods")

	sheets := f.GetSheetList()
	if len(sheets) != 5 {
		t.Fatalf("sheet count = %d, want 5", len(sheets))
	}

	expectedNames := []string{"Sheet1", "Sheet2", "Sheet3", "Sheet4", "Sheet5"}
	for i, name := range expectedNames {
		if sheets[i] != name {
			t.Errorf("sheet[%d] = %q, want %q", i, sheets[i], name)
		}
	}

	assertCellValue(t, f, "Sheet1", "A1", "Sheet1 Data")
	assertCellValue(t, f, "Sheet2", "A1", "Sheet2 Data")
	assertCellValue(t, f, "Sheet3", "A1", "Cross-Sheet Sums")

	f2 := saveAndReopen(t, f)

	sheets2 := f2.GetSheetList()
	if len(sheets2) != 5 {
		t.Fatalf("sheet count after roundtrip = %d, want 5", len(sheets2))
	}
	for i, name := range expectedNames {
		if sheets2[i] != name {
			t.Errorf("sheet[%d] after roundtrip = %q, want %q", i, sheets2[i], name)
		}
	}

	if err := f2.RecalcAll(); err != nil {
		t.Fatalf("RecalcAll: %v", err)
	}

	assertCellFloat(t, f2, "Sheet3", "A2", 550)
	assertCellFloat(t, f2, "Sheet3", "A3", 275)
	assertCellFloat(t, f2, "Sheet3", "A4", 15)
}

func TestGoldenDoubleRoundTrip(t *testing.T) {
	f := loadGolden(t, "golden_types.ods")

	f2 := saveAndReopen(t, f)
	f3 := saveAndReopen(t, f2)

	verifyTypesFile(t, f3)
}

func TestGoldenBufferVsFile(t *testing.T) {
	f := loadGolden(t, "golden_types.ods")

	dir, err := os.MkdirTemp("", "goods-golden-buf-*")
	if err != nil {
		t.Fatalf("create temp dir: %v", err)
	}
	t.Cleanup(func() { os.RemoveAll(dir) })

	filePath := filepath.Join(dir, "file.ods")
	if err := f.SaveAs(filePath); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}
	fileOpened, err := OpenFile(filePath)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	f2 := loadGolden(t, "golden_types.ods")
	buf, err := f2.WriteToBuffer()
	if err != nil {
		t.Fatalf("WriteToBuffer: %v", err)
	}
	bufOpened, err := OpenBytes(buf.Bytes())
	if err != nil {
		t.Fatalf("OpenBytes: %v", err)
	}

	cells := [][2]string{
		{"Sheet1", "B2"},
		{"Sheet1", "B4"},
		{"Sheet1", "B5"},
		{"Sheet1", "B6"},
		{"Sheet1", "B8"},
		{"Sheet1", "B10"},
		{"Sheet1", "B12"},
		{"Sheet1", "B15"},
		{"Sheet2", "A1"},
		{"Sheet2", "A2"},
	}

	for _, c := range cells {
		vFile, err := fileOpened.GetCellValue(c[0], c[1])
		if err != nil {
			t.Errorf("file GetCellValue(%s, %s): %v", c[0], c[1], err)
			continue
		}
		vBuf, err := bufOpened.GetCellValue(c[0], c[1])
		if err != nil {
			t.Errorf("buf GetCellValue(%s, %s): %v", c[0], c[1], err)
			continue
		}
		if vFile != vBuf {
			t.Errorf("mismatch at %s!%s: file=%q buf=%q", c[0], c[1], vFile, vBuf)
		}
	}

	for _, c := range cells {
		tFile, _ := fileOpened.GetCellType(c[0], c[1])
		tBuf, _ := bufOpened.GetCellType(c[0], c[1])
		if tFile != tBuf {
			t.Errorf("type mismatch at %s!%s: file=%d buf=%d", c[0], c[1], tFile, tBuf)
		}
	}

}
