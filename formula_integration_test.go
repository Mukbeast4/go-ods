package goods

import (
	"fmt"
	"os"
	"testing"
)

const testODSPath = "testdata/sample.ods"
const testSheetName = "Sheet1"

func skipIfNoTestFile(t *testing.T) {
	t.Helper()
	if _, err := os.Stat(testODSPath); os.IsNotExist(err) {
		t.Skipf("test file %s not found, skipping", testODSPath)
	}
}

func TestOpenRealODS(t *testing.T) {
	skipIfNoTestFile(t)
	f, err := OpenFile(testODSPath)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}
	defer f.Close()

	sheets := f.GetSheetList()
	if len(sheets) == 0 {
		t.Fatal("no sheets found")
	}
	t.Logf("Sheets: %v", sheets)

	found := false
	for _, s := range sheets {
		if s == testSheetName {
			found = true
		}
	}
	if !found {
		t.Fatalf("sheet %q not found in %v", testSheetName, sheets)
	}
}

func TestReadSheetData(t *testing.T) {
	skipIfNoTestFile(t)
	f, err := OpenFile(testODSPath)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}
	defer f.Close()

	dim, err := f.GetSheetDimension(testSheetName)
	if err != nil {
		t.Fatalf("GetSheetDimension: %v", err)
	}
	t.Logf("Dimension: %s", dim)

	rows, err := f.GetRows(testSheetName)
	if err != nil {
		t.Fatalf("GetRows: %v", err)
	}
	t.Logf("Total rows: %d", len(rows))

	for i, row := range rows {
		if i < 10 {
			t.Logf("Row %d: %v", i+1, row)
		}
	}
}

func TestExtractFormulas(t *testing.T) {
	skipIfNoTestFile(t)
	f, err := OpenFile(testODSPath)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}
	defer f.Close()

	formulas, err := f.GetSheetFormulas(testSheetName)
	if err != nil {
		t.Fatalf("GetSheetFormulas: %v", err)
	}

	if len(formulas) == 0 {
		t.Fatal("no formulas found")
	}
	t.Logf("Found %d formulas", len(formulas))

	for cellRef, formula := range formulas {
		t.Logf("  %s = %s", cellRef, formula)
	}
}

func TestRecalcFromODS(t *testing.T) {
	skipIfNoTestFile(t)
	f, err := OpenFile(testODSPath)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}
	defer f.Close()

	if err := f.RecalcSheet(testSheetName); err != nil {
		t.Fatalf("RecalcSheet: %v", err)
	}

	e2, _ := f.GetCellFloat(testSheetName, "E2")
	if e2 != 255 {
		t.Errorf("E2 = %v, want 255 (85+92+78)", e2)
	}

	f2, _ := f.GetCellFloat(testSheetName, "F2")
	if f2 != 85 {
		t.Errorf("F2 = %v, want 85 (255/3)", f2)
	}

	g2, _ := f.GetCellValue(testSheetName, "G2")
	if g2 != "Pass" {
		t.Errorf("G2 = %q, want \"Pass\"", g2)
	}

	g5, _ := f.GetCellValue(testSheetName, "G5")
	if g5 != "Fail" {
		t.Errorf("G5 = %q, want \"Fail\" (Diana avg=66.67)", g5)
	}
}

func TestAdaptFormula(t *testing.T) {
	tests := []struct {
		formula  string
		fromRow  int
		toRow    int
		expected string
	}{
		{"[.A6]", 6, 7, "[.A7]"},
		{"[.A6:.B6]", 6, 10, "[.A10:.B10]"},
		{"[.A6]+[.B6]", 6, 8, "[.A8]+[.B8]"},
		{"[.A1]", 6, 7, "[.A1]"},
	}

	for _, tt := range tests {
		t.Run(tt.formula, func(t *testing.T) {
			got := AdaptFormula(tt.formula, tt.fromRow, tt.toRow)
			if got != tt.expected {
				t.Errorf("AdaptFormula(%q, %d, %d) = %q, want %q", tt.formula, tt.fromRow, tt.toRow, got, tt.expected)
			}
		})
	}
}

func TestEvaluateBasic(t *testing.T) {
	tests := []struct {
		formula  string
		values   CellValues
		expected interface{}
	}{
		{"1+2", nil, 3.0},
		{"10*5", nil, 50.0},
		{"100/4", nil, 25.0},
		{"2+3*4", nil, 14.0},
		{"-5+10", nil, 5.0},
		{"(2+3)*4", nil, 20.0},
	}

	for _, tt := range tests {
		t.Run(tt.formula, func(t *testing.T) {
			got, err := Evaluate(tt.formula, tt.values)
			if err != nil {
				t.Fatalf("Evaluate(%q): %v", tt.formula, err)
			}
			if fmt.Sprintf("%v", got) != fmt.Sprintf("%v", tt.expected) {
				t.Errorf("Evaluate(%q) = %v, want %v", tt.formula, got, tt.expected)
			}
		})
	}
}

func TestEvaluateFunctions(t *testing.T) {
	tests := []struct {
		formula  string
		values   CellValues
		expected interface{}
	}{
		{"SUM(1;2;3)", nil, 6.0},
		{"MIN(5;3;8)", nil, 3.0},
		{"MAX(5;3;8)", nil, 8.0},
		{"ABS(-10)", nil, 10.0},
		{"ROUND(3.456;2)", nil, 3.46},
		{"IF(1>0;10;20)", nil, 10.0},
		{"IF(1<0;10;20)", nil, 20.0},
		{"AND(1;1)", nil, true},
		{"OR(0;1)", nil, true},
		{"NOT(0)", nil, true},
		{"CONCATENATE(\"hello\";\" \";\"world\")", nil, "hello world"},
		{"LEN(\"test\")", nil, 4.0},
		{"AVERAGE(10;20;30)", nil, 20.0},
		{"MOD(10;3)", nil, 1.0},
		{"POWER(2;10)", nil, 1024.0},
		{"SQRT(144)", nil, 12.0},
	}

	for _, tt := range tests {
		t.Run(tt.formula, func(t *testing.T) {
			got, err := Evaluate(tt.formula, tt.values)
			if err != nil {
				t.Fatalf("Evaluate(%q): %v", tt.formula, err)
			}
			if fmt.Sprintf("%v", got) != fmt.Sprintf("%v", tt.expected) {
				t.Errorf("Evaluate(%q) = %v, want %v", tt.formula, got, tt.expected)
			}
		})
	}
}

func TestEvaluateWithCellRefs(t *testing.T) {
	values := CellValues{
		"A6": 100.0,
		"B6": 200.0,
		"C6": 50.0,
	}

	tests := []struct {
		formula  string
		expected interface{}
	}{
		{"[.A6]+[.B6]", 300.0},
		{"[.A6]*2", 200.0},
		{"SUM([.A6];[.B6];[.C6])", 350.0},
		{"IF([.A6]>[.C6];[.A6];[.C6])", 100.0},
		{"[.A6]-[.C6]+[.B6]", 250.0},
	}

	for _, tt := range tests {
		t.Run(tt.formula, func(t *testing.T) {
			got, err := Evaluate(tt.formula, values)
			if err != nil {
				t.Fatalf("Evaluate(%q): %v", tt.formula, err)
			}
			if fmt.Sprintf("%v", got) != fmt.Sprintf("%v", tt.expected) {
				t.Errorf("Evaluate(%q) = %v, want %v", tt.formula, got, tt.expected)
			}
		})
	}
}

func TestEvaluateRange(t *testing.T) {
	values := CellValues{
		"A1": 10.0,
		"B1": 20.0,
		"C1": 30.0,
	}

	got, err := Evaluate("SUM([.A1:.C1])", values)
	if err != nil {
		t.Fatalf("Evaluate SUM range: %v", err)
	}
	if got != 60.0 {
		t.Errorf("SUM([.A1:.C1]) = %v, want 60", got)
	}
}

func TestEvaluateComparisons(t *testing.T) {
	tests := []struct {
		formula  string
		expected bool
	}{
		{"5>3", true},
		{"3>5", false},
		{"5=5", true},
		{"5<>3", true},
		{"5<=5", true},
		{"5>=6", false},
	}

	for _, tt := range tests {
		t.Run(tt.formula, func(t *testing.T) {
			got, err := Evaluate(tt.formula, nil)
			if err != nil {
				t.Fatalf("Evaluate(%q): %v", tt.formula, err)
			}
			if got != tt.expected {
				t.Errorf("Evaluate(%q) = %v, want %v", tt.formula, got, tt.expected)
			}
		})
	}
}

func TestFormulaToGoBasic(t *testing.T) {
	tests := []struct {
		formula  string
		expected string
	}{
		{"[.A6]+[.B6]", "r.A+r.B"},
		{"[.A6]*2", "r.A*2"},
		{"[.A6]=[.B6]", "r.A==r.B"},
		{"[.A6]<>[.B6]", "r.A!=r.B"},
	}

	for _, tt := range tests {
		t.Run(tt.formula, func(t *testing.T) {
			got := FormulaToGo(tt.formula)
			if got != tt.expected {
				t.Errorf("FormulaToGo(%q) = %q, want %q", tt.formula, got, tt.expected)
			}
		})
	}
}

func TestCleanFormulaBasic(t *testing.T) {
	tests := []struct {
		input    string
		expected string
	}{
		{"of:=[.A6]+[.B6]", "A6+B6"},
		{"of:=SUM([.A1:.C1])", "SUM(A1:C1)"},
		{"of:=IF([.A6]&gt;0;[.B6];[.C6])", "IF(A6>0, B6, C6)"},
	}

	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			got := CleanFormula(tt.input)
			if got != tt.expected {
				t.Errorf("CleanFormula(%q) = %q, want %q", tt.input, got, tt.expected)
			}
		})
	}
}

func TestCopyRowFormulas(t *testing.T) {
	f := NewFile()
	sheet := "Sheet1"

	f.SetCellFloat(sheet, "A2", 10)
	f.SetCellFloat(sheet, "B2", 20)
	f.SetCellFormula(sheet, "C2", "[.A2]+[.B2]")
	f.SetCellFormula(sheet, "D2", "[.A2]*[.B2]")

	err := f.CopyRowFormulas(sheet, 2, 5)
	if err != nil {
		t.Fatalf("CopyRowFormulas: %v", err)
	}

	c5, _ := f.GetCellFormula(sheet, "C5")
	d5, _ := f.GetCellFormula(sheet, "D5")

	expectedC := "[.A5]+[.B5]"
	expectedD := "[.A5]*[.B5]"

	if c5 != expectedC {
		t.Errorf("C5 formula = %q, want %q", c5, expectedC)
	}
	if d5 != expectedD {
		t.Errorf("D5 formula = %q, want %q", d5, expectedD)
	}
}
