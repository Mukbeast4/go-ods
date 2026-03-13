package goods

import (
	"testing"
)

func TestCompareCleanFormulaOutput(t *testing.T) {
	tests := []struct {
		odsRaw   string
		expected string
	}{
		{
			"of:=[.A6]+[.B6]",
			"A6+B6",
		},
		{
			"of:=SUM([.R6:.X6])+([.J6]+[.K6]+[.L6]+[.M6])/3",
			"SUM(R6:X6)+(J6+K6+L6+M6)/3",
		},
		{
			"of:=IF([.A6]&gt;0;[.B6];[.C6])",
			"IF(A6>0, B6, C6)",
		},
		{
			"of:=IF(AND([.E6]=\"\";[.G6]+[.H6]<22;[.AG6]=0);\"BUY\";\"\")",
			"IF(AND(E6=\"\", G6+H6<22, AG6=0), \"BUY\", \"\")",
		},
	}

	for _, tt := range tests {
		got := CleanFormula(tt.odsRaw)
		if got != tt.expected {
			t.Errorf("CleanFormula(%q)\n  got:  %s\n  want: %s", tt.odsRaw, got, tt.expected)
		}
	}
}

func TestCompareToGoOutput(t *testing.T) {
	tests := []struct {
		odsRaw   string
		expected string
	}{
		{
			"[.A6]+[.B6]",
			"r.A+r.B",
		},
		{
			"[.AB6]+[.AE6]",
			"r.AB+r.AE",
		},
		{
			"IF(AND([.E6]=\"\";[.G6]>11);\"BUY\";\"\")",
			"func() interface{} { if (r.E==\"\" && r.G>11) { return \"BUY\" } else { return \"\" } }()",
		},
	}

	for _, tt := range tests {
		got := FormulaToGo(tt.odsRaw)
		if got != tt.expected {
			t.Errorf("FormulaToGo(%q)\n  got:  %s\n  want: %s", tt.odsRaw, got, tt.expected)
		}
	}
}

func TestCompareEvaluationGeneric(t *testing.T) {
	f := NewFile()
	sheet := "Sheet1"

	f.SetCellFloat(sheet, "A1", 100)
	f.SetCellFloat(sheet, "B1", 200)
	f.SetCellFloat(sheet, "C1", 50)
	f.SetCellFormula(sheet, "D1", "[.A1]+[.B1]+[.C1]")
	f.SetCellFormula(sheet, "E1", "([.A1]+[.B1]+[.C1])/3")

	f.RecalcSheet(sheet)

	d1, _ := f.GetCellFloat(sheet, "D1")
	e1, _ := f.GetCellFloat(sheet, "E1")

	if d1 != 350 {
		t.Errorf("D1 = %v, want 350", d1)
	}
	expected := (100.0 + 200.0 + 50.0) / 3.0
	if e1 != expected {
		t.Errorf("E1 = %v, want %v", e1, expected)
	}
}

func getCellByRef(s *sheet, ref string) *cell {
	col, row, err := CellNameToCoordinates(ref)
	if err != nil {
		return nil
	}
	return s.getCell(col, row)
}
