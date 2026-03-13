package goods

import (
	"fmt"
	"testing"
)

func TestCompareEvalWithStoredValues(t *testing.T) {
	f := NewFile()
	sheet := "Sheet1"

	f.SetCellFloat(sheet, "A1", 10)
	f.SetCellFloat(sheet, "B1", 20)
	f.SetCellFloat(sheet, "C1", 30)
	f.SetCellFloat(sheet, "D1", 1)

	f.SetCellFormula(sheet, "E1", "SUM([.A1:.C1])")
	f.SetCellFormula(sheet, "F1", "([.A1]+[.B1]+[.C1])/3")
	f.SetCellFormula(sheet, "G1", "IF([.D1]=1;[.A1];0)")

	f.RecalcSheet(sheet)

	tests := []struct {
		ref      string
		expected float64
	}{
		{"E1", 60},
		{"F1", 20},
		{"G1", 10},
	}

	for _, tt := range tests {
		val, _ := f.GetCellFloat(sheet, tt.ref)
		if val != tt.expected {
			t.Errorf("%s = %v, want %v", tt.ref, val, tt.expected)
		}
	}
}

func TestCompareChainedEval(t *testing.T) {
	f := NewFile()
	sheet := "Sheet1"

	f.SetCellFloat(sheet, "A1", 5)
	f.SetCellFloat(sheet, "B1", 3)
	f.SetCellFormula(sheet, "C1", "[.A1]*[.B1]")
	f.SetCellFormula(sheet, "D1", "[.C1]+10")
	f.SetCellFormula(sheet, "E1", "IF([.D1]>20;\"high\";\"low\")")

	f.RecalcSheet(sheet)

	c1, _ := f.GetCellFloat(sheet, "C1")
	d1, _ := f.GetCellFloat(sheet, "D1")
	e1, _ := f.GetCellValue(sheet, "E1")

	if c1 != 15 {
		t.Errorf("C1 = %v, want 15", c1)
	}
	if d1 != 25 {
		t.Errorf("D1 = %v, want 25", d1)
	}
	if e1 != "high" {
		t.Errorf("E1 = %q, want \"high\"", e1)
	}
}

func val(v CellValues, key string) float64 {
	if f, ok := toFloat(v[key]); ok {
		return f
	}
	return 0
}

func TestManualCalculation(t *testing.T) {
	values := CellValues{
		"A1": 10.0,
		"B1": 20.0,
		"C1": 30.0,
	}

	sum := val(values, "A1") + val(values, "B1") + val(values, "C1")
	avg := sum / 3

	result, err := Evaluate("SUM([.A1];[.B1];[.C1])", values)
	if err != nil {
		t.Fatalf("Evaluate SUM: %v", err)
	}
	if result != sum {
		t.Errorf("SUM = %v, want %v", result, sum)
	}

	result2, err := Evaluate("([.A1]+[.B1]+[.C1])/3", values)
	if err != nil {
		t.Fatalf("Evaluate AVG: %v", err)
	}
	if fmt.Sprintf("%.2f", result2) != fmt.Sprintf("%.2f", avg) {
		t.Errorf("AVG = %v, want %v", result2, avg)
	}
}
