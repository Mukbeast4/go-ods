package goods

import (
	"testing"
)

func TestRecalcSimpleChain(t *testing.T) {
	f := NewFile()
	sheet := "Sheet1"

	f.SetCellFloat(sheet, "A1", 10)
	f.SetCellFormula(sheet, "B1", "[.A1]*2")
	f.SetCellFormula(sheet, "C1", "[.B1]+5")

	if err := f.RecalcSheet(sheet); err != nil {
		t.Fatalf("RecalcSheet: %v", err)
	}

	b1, _ := f.GetCellFloat(sheet, "B1")
	c1, _ := f.GetCellFloat(sheet, "C1")

	if b1 != 20 {
		t.Errorf("B1 = %v, want 20", b1)
	}
	if c1 != 25 {
		t.Errorf("C1 = %v, want 25", c1)
	}
}

func TestRecalcAfterValueChange(t *testing.T) {
	f := NewFile()
	sheet := "Sheet1"

	f.SetCellFloat(sheet, "A1", 10)
	f.SetCellFormula(sheet, "B1", "[.A1]*2")
	f.SetCellFormula(sheet, "C1", "[.B1]+5")

	f.RecalcSheet(sheet)

	f.SetCellFloat(sheet, "A1", 20)
	f.RecalcSheet(sheet)

	b1, _ := f.GetCellFloat(sheet, "B1")
	c1, _ := f.GetCellFloat(sheet, "C1")

	if b1 != 40 {
		t.Errorf("B1 = %v, want 40", b1)
	}
	if c1 != 45 {
		t.Errorf("C1 = %v, want 45", c1)
	}
}

func TestRecalcCircularReference(t *testing.T) {
	f := NewFile()
	sheet := "Sheet1"

	f.SetCellFormula(sheet, "A1", "[.B1]+1")
	f.SetCellFormula(sheet, "B1", "[.A1]+1")

	err := f.RecalcSheet(sheet)
	if err != ErrCircularReference {
		t.Errorf("expected ErrCircularReference, got %v", err)
	}
}

func TestRecalcAutoRecalc(t *testing.T) {
	f := NewFile()
	sheet := "Sheet1"

	f.SetCellFloat(sheet, "A1", 10)
	f.SetCellFormula(sheet, "B1", "[.A1]*3")

	f.SetAutoRecalc(true)

	f.SetCellFloat(sheet, "A1", 5)

	b1, _ := f.GetCellFloat(sheet, "B1")
	if b1 != 15 {
		t.Errorf("B1 = %v, want 15", b1)
	}

	f.SetCellFloat(sheet, "A1", 100)

	b1, _ = f.GetCellFloat(sheet, "B1")
	if b1 != 300 {
		t.Errorf("B1 = %v, want 300", b1)
	}
}

func TestRecalcRange(t *testing.T) {
	f := NewFile()
	sheet := "Sheet1"

	f.SetCellFloat(sheet, "A1", 10)
	f.SetCellFloat(sheet, "A2", 20)
	f.SetCellFloat(sheet, "A3", 30)
	f.SetCellFormula(sheet, "A4", "SUM([.A1:.A3])")

	if err := f.RecalcSheet(sheet); err != nil {
		t.Fatalf("RecalcSheet: %v", err)
	}

	a4, _ := f.GetCellFloat(sheet, "A4")
	if a4 != 60 {
		t.Errorf("A4 = %v, want 60", a4)
	}
}

func TestRecalcNoFormulas(t *testing.T) {
	f := NewFile()
	sheet := "Sheet1"

	f.SetCellFloat(sheet, "A1", 10)
	f.SetCellStr(sheet, "B1", "hello")

	if err := f.RecalcSheet(sheet); err != nil {
		t.Errorf("RecalcSheet with no formulas should not error, got %v", err)
	}
}

func TestRecalcLongChain(t *testing.T) {
	f := NewFile()
	sheet := "Sheet1"

	f.SetCellFloat(sheet, "A1", 1)
	f.SetCellFormula(sheet, "B1", "[.A1]+1")
	f.SetCellFormula(sheet, "C1", "[.B1]+1")
	f.SetCellFormula(sheet, "D1", "[.C1]+1")

	if err := f.RecalcSheet(sheet); err != nil {
		t.Fatalf("RecalcSheet: %v", err)
	}

	d1, _ := f.GetCellFloat(sheet, "D1")
	if d1 != 4 {
		t.Errorf("D1 = %v, want 4", d1)
	}

	f.SetCellFloat(sheet, "A1", 10)
	f.RecalcSheet(sheet)

	d1, _ = f.GetCellFloat(sheet, "D1")
	if d1 != 13 {
		t.Errorf("D1 = %v, want 13", d1)
	}
}

func TestRecalcMultipleFormulasOnSameSource(t *testing.T) {
	f := NewFile()
	sheet := "Sheet1"

	f.SetCellFloat(sheet, "A1", 10)
	f.SetCellFormula(sheet, "B1", "[.A1]*2")
	f.SetCellFormula(sheet, "C1", "[.A1]*3")
	f.SetCellFormula(sheet, "D1", "[.B1]+[.C1]")

	if err := f.RecalcSheet(sheet); err != nil {
		t.Fatalf("RecalcSheet: %v", err)
	}

	b1, _ := f.GetCellFloat(sheet, "B1")
	c1, _ := f.GetCellFloat(sheet, "C1")
	d1, _ := f.GetCellFloat(sheet, "D1")

	if b1 != 20 {
		t.Errorf("B1 = %v, want 20", b1)
	}
	if c1 != 30 {
		t.Errorf("C1 = %v, want 30", c1)
	}
	if d1 != 50 {
		t.Errorf("D1 = %v, want 50", d1)
	}
}

func TestRecalcAll(t *testing.T) {
	f := NewFile()
	f.NewSheet("Sheet2")

	f.SetCellFloat("Sheet1", "A1", 5)
	f.SetCellFormula("Sheet1", "B1", "[.A1]*2")

	f.SetCellFloat("Sheet2", "A1", 100)
	f.SetCellFormula("Sheet2", "B1", "[.A1]/2")

	if err := f.RecalcAll(); err != nil {
		t.Fatalf("RecalcAll: %v", err)
	}

	b1s1, _ := f.GetCellFloat("Sheet1", "B1")
	b1s2, _ := f.GetCellFloat("Sheet2", "B1")

	if b1s1 != 10 {
		t.Errorf("Sheet1 B1 = %v, want 10", b1s1)
	}
	if b1s2 != 50 {
		t.Errorf("Sheet2 B1 = %v, want 50", b1s2)
	}
}

func TestRecalcWithIF(t *testing.T) {
	f := NewFile()
	sheet := "Sheet1"

	f.SetCellFloat(sheet, "A1", 10)
	f.SetCellFormula(sheet, "B1", "IF([.A1]>5;\"big\";\"small\")")

	if err := f.RecalcSheet(sheet); err != nil {
		t.Fatalf("RecalcSheet: %v", err)
	}

	b1, _ := f.GetCellValue(sheet, "B1")
	if b1 != "big" {
		t.Errorf("B1 = %q, want \"big\"", b1)
	}

	f.SetCellFloat(sheet, "A1", 3)
	f.RecalcSheet(sheet)

	b1, _ = f.GetCellValue(sheet, "B1")
	if b1 != "small" {
		t.Errorf("B1 = %q, want \"small\"", b1)
	}
}

func TestExtractRefs(t *testing.T) {
	tests := []struct {
		formula  string
		expected int
	}{
		{"[.A1]+[.B1]", 2},
		{"SUM([.A1:.A3])", 3},
		{"[.A1]*2+[.B1]*3", 2},
		{"IF([.A1]>5;[.B1];[.C1])", 3},
		{"10+20", 0},
		{"\"[.A1]\"", 0},
	}

	for _, tt := range tests {
		refs := extractRefs(tt.formula)
		if len(refs) != tt.expected {
			t.Errorf("extractRefs(%q) = %d refs, want %d", tt.formula, len(refs), tt.expected)
		}
	}
}

func TestRecalcSelfReference(t *testing.T) {
	f := NewFile()
	sheet := "Sheet1"

	f.SetCellFormula(sheet, "A1", "[.A1]+1")

	err := f.RecalcSheet(sheet)
	if err != ErrCircularReference {
		t.Errorf("expected ErrCircularReference for self-reference, got %v", err)
	}
}

func TestRecalcClosedFile(t *testing.T) {
	f := NewFile()
	f.Close()

	err := f.RecalcSheet("Sheet1")
	if err != ErrFileClosed {
		t.Errorf("expected ErrFileClosed, got %v", err)
	}

	err = f.RecalcAll()
	if err != ErrFileClosed {
		t.Errorf("expected ErrFileClosed, got %v", err)
	}
}

func TestRecalcSheetNotFound(t *testing.T) {
	f := NewFile()

	err := f.RecalcSheet("NonExistent")
	if err != ErrSheetNotFound {
		t.Errorf("expected ErrSheetNotFound, got %v", err)
	}
}
