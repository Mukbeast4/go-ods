package goods

import (
	"math"
	"testing"
)

func TestVLOOKUP(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellStr(s, "A1", "Apple")
	f.SetCellFloat(s, "B1", 1.5)
	f.SetCellStr(s, "A2", "Banana")
	f.SetCellFloat(s, "B2", 2.0)
	f.SetCellStr(s, "A3", "Cherry")
	f.SetCellFloat(s, "B3", 3.5)

	f.SetCellFormula(s, "D1", "VLOOKUP(\"Banana\";[.A1:.B3];2;0)")
	f.RecalcSheet(s)

	v, _ := f.GetCellFloat(s, "D1")
	if v != 2.0 {
		t.Errorf("VLOOKUP got %v, want 2.0", v)
	}
}

func TestVLOOKUPNotFound(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellStr(s, "A1", "Apple")
	f.SetCellFloat(s, "B1", 1.5)
	f.SetCellFormula(s, "D1", "VLOOKUP(\"Mango\";[.A1:.B1];2;0)")

	err := f.RecalcSheet(s)
	if err != nil {
		t.Fatalf("RecalcSheet: %v", err)
	}
}

func TestHLOOKUP(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellStr(s, "A1", "Q1")
	f.SetCellStr(s, "B1", "Q2")
	f.SetCellStr(s, "C1", "Q3")
	f.SetCellFloat(s, "A2", 100)
	f.SetCellFloat(s, "B2", 200)
	f.SetCellFloat(s, "C2", 300)

	f.SetCellFormula(s, "E1", "HLOOKUP(\"Q2\";[.A1:.C2];2;0)")
	f.RecalcSheet(s)

	v, _ := f.GetCellFloat(s, "E1")
	if v != 200 {
		t.Errorf("HLOOKUP got %v, want 200", v)
	}
}

func TestINDEX(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellFloat(s, "A1", 10)
	f.SetCellFloat(s, "B1", 20)
	f.SetCellFloat(s, "A2", 30)
	f.SetCellFloat(s, "B2", 40)

	f.SetCellFormula(s, "D1", "INDEX([.A1:.B2];2;2)")
	f.RecalcSheet(s)

	v, _ := f.GetCellFloat(s, "D1")
	if v != 40 {
		t.Errorf("INDEX got %v, want 40", v)
	}
}

func TestMATCH(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellFloat(s, "A1", 10)
	f.SetCellFloat(s, "A2", 20)
	f.SetCellFloat(s, "A3", 30)

	f.SetCellFormula(s, "B1", "MATCH(20;[.A1:.A3])")
	f.RecalcSheet(s)

	v, _ := f.GetCellFloat(s, "B1")
	if v != 2 {
		t.Errorf("MATCH got %v, want 2", v)
	}
}

func TestINDEXMATCH(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellStr(s, "A1", "X")
	f.SetCellStr(s, "A2", "Y")
	f.SetCellStr(s, "A3", "Z")
	f.SetCellFloat(s, "B1", 100)
	f.SetCellFloat(s, "B2", 200)
	f.SetCellFloat(s, "B3", 300)

	f.SetCellFormula(s, "D1", "INDEX([.B1:.B3];MATCH(\"Y\";[.A1:.A3]))")
	f.RecalcSheet(s)

	v, _ := f.GetCellFloat(s, "D1")
	if v != 200 {
		t.Errorf("INDEX/MATCH got %v, want 200", v)
	}
}

func TestIFERROR(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellFormula(s, "A1", "IFERROR(1/0;\"err\")")
	f.RecalcSheet(s)

	v, _ := f.GetCellValue(s, "A1")
	if v != "err" {
		t.Errorf("IFERROR got %q, want \"err\"", v)
	}
}

func TestIFERRORNoError(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellFloat(s, "B1", 5)
	f.SetCellFormula(s, "A1", "IFERROR([.B1]*2;\"err\")")
	f.RecalcSheet(s)

	v, _ := f.GetCellFloat(s, "A1")
	if v != 10 {
		t.Errorf("IFERROR got %v, want 10", v)
	}
}

func TestSUMIF(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellStr(s, "A1", "fruit")
	f.SetCellStr(s, "A2", "veggie")
	f.SetCellStr(s, "A3", "fruit")
	f.SetCellFloat(s, "B1", 10)
	f.SetCellFloat(s, "B2", 20)
	f.SetCellFloat(s, "B3", 30)

	f.SetCellFormula(s, "C1", "SUMIF([.A1:.A3];\"fruit\";[.B1:.B3])")
	f.RecalcSheet(s)

	v, _ := f.GetCellFloat(s, "C1")
	if v != 40 {
		t.Errorf("SUMIF got %v, want 40", v)
	}
}

func TestSUMIFNumericCriteria(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellFloat(s, "A1", 10)
	f.SetCellFloat(s, "A2", 20)
	f.SetCellFloat(s, "A3", 30)

	f.SetCellFormula(s, "B1", "SUMIF([.A1:.A3];\">15\")")
	f.RecalcSheet(s)

	v, _ := f.GetCellFloat(s, "B1")
	if v != 50 {
		t.Errorf("SUMIF got %v, want 50", v)
	}
}

func TestCOUNTIF(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellFloat(s, "A1", 1)
	f.SetCellFloat(s, "A2", 2)
	f.SetCellFloat(s, "A3", 3)
	f.SetCellFloat(s, "A4", 2)

	f.SetCellFormula(s, "B1", "COUNTIF([.A1:.A4];\"2\")")
	f.RecalcSheet(s)

	v, _ := f.GetCellFloat(s, "B1")
	if v != 2 {
		t.Errorf("COUNTIF got %v, want 2", v)
	}
}

func TestCOUNTA(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellStr(s, "A1", "hello")
	f.SetCellFloat(s, "A2", 5)
	f.SetCellStr(s, "A3", "world")

	f.SetCellFormula(s, "B1", "COUNTA([.A1:.A3])")
	f.RecalcSheet(s)

	v, _ := f.GetCellFloat(s, "B1")
	if v != 3 {
		t.Errorf("COUNTA got %v, want 3", v)
	}
}

func TestDATE(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellFormula(s, "A1", "DATE(2024;1;15)")
	f.RecalcSheet(s)

	v, _ := f.GetCellFloat(s, "A1")
	expected := 45306.0
	if math.Abs(v-expected) > 1 {
		t.Errorf("DATE got %v, want ~%v", v, expected)
	}
}

func TestYEAR(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellFormula(s, "A1", "DATE(2024;6;15)")
	f.SetCellFormula(s, "B1", "YEAR([.A1])")
	f.RecalcSheet(s)

	v, _ := f.GetCellFloat(s, "B1")
	if v != 2024 {
		t.Errorf("YEAR got %v, want 2024", v)
	}
}

func TestMONTH(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellFormula(s, "A1", "DATE(2024;6;15)")
	f.SetCellFormula(s, "B1", "MONTH([.A1])")
	f.RecalcSheet(s)

	v, _ := f.GetCellFloat(s, "B1")
	if v != 6 {
		t.Errorf("MONTH got %v, want 6", v)
	}
}

func TestDAY(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellFormula(s, "A1", "DATE(2024;6;15)")
	f.SetCellFormula(s, "B1", "DAY([.A1])")
	f.RecalcSheet(s)

	v, _ := f.GetCellFloat(s, "B1")
	if v != 15 {
		t.Errorf("DAY got %v, want 15", v)
	}
}

func TestFIND(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellStr(s, "A1", "Hello World")
	f.SetCellFormula(s, "B1", "FIND(\"World\";[.A1])")
	f.RecalcSheet(s)

	v, _ := f.GetCellFloat(s, "B1")
	if v != 7 {
		t.Errorf("FIND got %v, want 7", v)
	}
}

func TestFINDCaseSensitive(t *testing.T) {
	result, err := Evaluate("FIND(\"world\";\"Hello World\")", CellValues{})
	if err == nil {
		t.Errorf("FIND should be case-sensitive, got %v", result)
	}
}

func TestSEARCH(t *testing.T) {
	result, err := Evaluate("SEARCH(\"world\";\"Hello World\")", CellValues{})
	if err != nil {
		t.Fatalf("SEARCH: %v", err)
	}
	v, _ := toFloat(result)
	if v != 7 {
		t.Errorf("SEARCH got %v, want 7", v)
	}
}

func TestSUBSTITUTE(t *testing.T) {
	result, err := Evaluate("SUBSTITUTE(\"aaa\";\"a\";\"b\")", CellValues{})
	if err != nil {
		t.Fatalf("SUBSTITUTE: %v", err)
	}
	if result != "bbb" {
		t.Errorf("SUBSTITUTE got %v, want bbb", result)
	}
}

func TestSUBSTITUTEInstance(t *testing.T) {
	result, err := Evaluate("SUBSTITUTE(\"aaa\";\"a\";\"b\";2)", CellValues{})
	if err != nil {
		t.Fatalf("SUBSTITUTE: %v", err)
	}
	if result != "aba" {
		t.Errorf("SUBSTITUTE got %v, want aba", result)
	}
}

func TestREPLACE(t *testing.T) {
	result, err := Evaluate("REPLACE(\"Hello\";2;3;\"XY\")", CellValues{})
	if err != nil {
		t.Fatalf("REPLACE: %v", err)
	}
	if result != "HXYo" {
		t.Errorf("REPLACE got %v, want HXYo", result)
	}
}

func TestTEXT(t *testing.T) {
	result, err := Evaluate("TEXT(0.75;\"0%\")", CellValues{})
	if err != nil {
		t.Fatalf("TEXT: %v", err)
	}
	if result != "75%" {
		t.Errorf("TEXT got %v, want 75%%", result)
	}
}

func TestVALUE(t *testing.T) {
	result, err := Evaluate("VALUE(\"42.5\")", CellValues{})
	if err != nil {
		t.Fatalf("VALUE: %v", err)
	}
	v, _ := toFloat(result)
	if v != 42.5 {
		t.Errorf("VALUE got %v, want 42.5", v)
	}
}

func TestINT(t *testing.T) {
	result, err := Evaluate("INT(3.7)", CellValues{})
	if err != nil {
		t.Fatalf("INT: %v", err)
	}
	v, _ := toFloat(result)
	if v != 3 {
		t.Errorf("INT got %v, want 3", v)
	}
}

func TestINTNegative(t *testing.T) {
	result, err := Evaluate("INT(-3.2)", CellValues{})
	if err != nil {
		t.Fatalf("INT: %v", err)
	}
	v, _ := toFloat(result)
	if v != -4 {
		t.Errorf("INT got %v, want -4", v)
	}
}

func TestSUMPRODUCT(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellFloat(s, "A1", 1)
	f.SetCellFloat(s, "A2", 2)
	f.SetCellFloat(s, "A3", 3)
	f.SetCellFloat(s, "B1", 4)
	f.SetCellFloat(s, "B2", 5)
	f.SetCellFloat(s, "B3", 6)

	f.SetCellFormula(s, "C1", "SUMPRODUCT([.A1:.A3];[.B1:.B3])")
	f.RecalcSheet(s)

	v, _ := f.GetCellFloat(s, "C1")
	if v != 32 {
		t.Errorf("SUMPRODUCT got %v, want 32", v)
	}
}

func TestTODAY(t *testing.T) {
	result, err := Evaluate("TODAY()", CellValues{})
	if err != nil {
		t.Fatalf("TODAY: %v", err)
	}
	v, ok := toFloat(result)
	if !ok || v < 40000 {
		t.Errorf("TODAY returned unexpected value: %v", result)
	}
}

func TestCrossSheetRef(t *testing.T) {
	f := NewFile()
	f.NewSheet("Data")

	f.SetCellFloat("Data", "A1", 42)
	f.SetCellFormula("Sheet1", "A1", "[Data.A1]*2")

	err := f.RecalcSheet("Sheet1")
	if err != nil {
		t.Fatalf("RecalcSheet: %v", err)
	}

	v, _ := f.GetCellFloat("Sheet1", "A1")
	if v != 84 {
		t.Errorf("cross-sheet ref got %v, want 84", v)
	}
}

func TestCrossSheetRange(t *testing.T) {
	f := NewFile()
	f.NewSheet("Data")

	f.SetCellFloat("Data", "A1", 10)
	f.SetCellFloat("Data", "A2", 20)
	f.SetCellFloat("Data", "A3", 30)
	f.SetCellFormula("Sheet1", "A1", "SUM([Data.A1:.A3])")

	err := f.RecalcSheet("Sheet1")
	if err != nil {
		t.Fatalf("RecalcSheet: %v", err)
	}

	v, _ := f.GetCellFloat("Sheet1", "A1")
	if v != 60 {
		t.Errorf("cross-sheet range got %v, want 60", v)
	}
}

func TestEmptyCellNotEqualEmptyString(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellFormula(s, "A1", "IF([.B1]<>\"\";\"YES\";\"NO\")")
	f.RecalcSheet(s)

	v, _ := f.GetCellValue(s, "A1")
	if v != "NO" {
		t.Errorf("empty cell <> \"\" should be false, got %q", v)
	}
}

func TestEmptyCellsInOR(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellFormula(s, "D1", "IF(OR([.A1]<>\"\",[.B1]<>\"\",[.C1]<>\"\");\"YES\";\"\")")
	f.RecalcSheet(s)

	v, _ := f.GetCellValue(s, "D1")
	if v != "" {
		t.Errorf("OR with all empty cells should return \"\", got %q", v)
	}
}

func TestEmptyCellsInORWithValue(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellStr(s, "B1", "hello")
	f.SetCellFormula(s, "D1", "IF(OR([.A1]<>\"\",[.B1]<>\"\",[.C1]<>\"\");\"YES\";\"\")")
	f.RecalcSheet(s)

	v, _ := f.GetCellValue(s, "D1")
	if v != "YES" {
		t.Errorf("OR with one non-empty cell should return YES, got %q", v)
	}
}

func TestEmptyCellArithmetic(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellFloat(s, "A1", 5)
	f.SetCellFormula(s, "C1", "[.A1]+[.B1]")
	f.RecalcSheet(s)

	v, _ := f.GetCellFloat(s, "C1")
	if v != 5 {
		t.Errorf("5 + empty should be 5, got %v", v)
	}
}

func TestCountSkipsEmpty(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellFloat(s, "A1", 10)
	f.SetCellFloat(s, "A3", 30)
	f.SetCellFormula(s, "B1", "COUNT([.A1:.A3])")
	f.RecalcSheet(s)

	v, _ := f.GetCellFloat(s, "B1")
	if v != 2 {
		t.Errorf("COUNT should skip empty cells, got %v", v)
	}
}

func TestAverageSkipsEmpty(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellFloat(s, "A1", 10)
	f.SetCellFloat(s, "A3", 30)
	f.SetCellFormula(s, "B1", "AVERAGE([.A1:.A3])")
	f.RecalcSheet(s)

	v, _ := f.GetCellFloat(s, "B1")
	if v != 20 {
		t.Errorf("AVERAGE should skip empty cells, got %v", v)
	}
}

func TestCOUNTIFS(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellStr(s, "A1", "fruit")
	f.SetCellStr(s, "A2", "veggie")
	f.SetCellStr(s, "A3", "fruit")
	f.SetCellStr(s, "A4", "fruit")
	f.SetCellStr(s, "B1", "red")
	f.SetCellStr(s, "B2", "green")
	f.SetCellStr(s, "B3", "red")
	f.SetCellStr(s, "B4", "yellow")

	f.SetCellFormula(s, "C1", "COUNTIFS([.A1:.A4];\"fruit\";[.B1:.B4];\"red\")")
	f.RecalcSheet(s)

	v, _ := f.GetCellFloat(s, "C1")
	if v != 2 {
		t.Errorf("COUNTIFS got %v, want 2", v)
	}
}

func TestCOUNTIFSMultiple(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellStr(s, "A1", "a")
	f.SetCellStr(s, "A2", "a")
	f.SetCellStr(s, "A3", "b")
	f.SetCellStr(s, "B1", "x")
	f.SetCellStr(s, "B2", "x")
	f.SetCellStr(s, "B3", "x")
	f.SetCellFloat(s, "C1", 10)
	f.SetCellFloat(s, "C2", 20)
	f.SetCellFloat(s, "C3", 10)

	f.SetCellFormula(s, "D1", "COUNTIFS([.A1:.A3];\"a\";[.B1:.B3];\"x\";[.C1:.C3];\"10\")")
	f.RecalcSheet(s)

	v, _ := f.GetCellFloat(s, "D1")
	if v != 1 {
		t.Errorf("COUNTIFS 3 criteria got %v, want 1", v)
	}
}

func TestSUMIFS(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellStr(s, "A1", "fruit")
	f.SetCellStr(s, "A2", "veggie")
	f.SetCellStr(s, "A3", "fruit")
	f.SetCellStr(s, "B1", "red")
	f.SetCellStr(s, "B2", "green")
	f.SetCellStr(s, "B3", "red")
	f.SetCellFloat(s, "C1", 10)
	f.SetCellFloat(s, "C2", 20)
	f.SetCellFloat(s, "C3", 30)

	f.SetCellFormula(s, "D1", "SUMIFS([.C1:.C3];[.A1:.A3];\"fruit\";[.B1:.B3];\"red\")")
	f.RecalcSheet(s)

	v, _ := f.GetCellFloat(s, "D1")
	if v != 40 {
		t.Errorf("SUMIFS got %v, want 40", v)
	}
}

func TestSUMIFSNumericCriteria(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellFloat(s, "A1", 5)
	f.SetCellFloat(s, "A2", 15)
	f.SetCellFloat(s, "A3", 25)
	f.SetCellFloat(s, "A4", 35)
	f.SetCellFloat(s, "B1", 100)
	f.SetCellFloat(s, "B2", 200)
	f.SetCellFloat(s, "B3", 300)
	f.SetCellFloat(s, "B4", 400)

	f.SetCellFormula(s, "C1", "SUMIFS([.B1:.B4];[.A1:.A4];\">10\";[.A1:.A4];\"<=30\")")
	f.RecalcSheet(s)

	v, _ := f.GetCellFloat(s, "C1")
	if v != 500 {
		t.Errorf("SUMIFS numeric got %v, want 500", v)
	}
}

func TestCOUNTIFSEmpty(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellFormula(s, "A1", "COUNTIFS([.B1:.B3];\"x\";[.C1:.C3];\"y\")")
	f.RecalcSheet(s)

	v, _ := f.GetCellFloat(s, "A1")
	if v != 0 {
		t.Errorf("COUNTIFS empty got %v, want 0", v)
	}
}

func TestSUMIFSNoMatch(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellStr(s, "A1", "a")
	f.SetCellStr(s, "A2", "b")
	f.SetCellFloat(s, "B1", 10)
	f.SetCellFloat(s, "B2", 20)

	f.SetCellFormula(s, "C1", "SUMIFS([.B1:.B2];[.A1:.A2];\"z\")")
	f.RecalcSheet(s)

	v, _ := f.GetCellFloat(s, "C1")
	if v != 0 {
		t.Errorf("SUMIFS no match got %v, want 0", v)
	}
}

func TestCrossSheetDepGraph(t *testing.T) {
	f := NewFile()
	f.NewSheet("Data")

	f.SetCellFloat("Data", "A1", 10)
	f.SetCellFloat("Data", "A2", 20)
	f.SetCellFormula("Sheet1", "A1", "[Data.A1]+[Data.A2]")

	err := f.RecalcAll()
	if err != nil {
		t.Fatalf("RecalcAll: %v", err)
	}

	v, _ := f.GetCellFloat("Sheet1", "A1")
	if v != 30 {
		t.Errorf("cross-sheet dep got %v, want 30", v)
	}
}

func TestCrossSheetDepChain(t *testing.T) {
	f := NewFile()
	f.NewSheet("Data")

	f.SetCellFloat("Data", "A1", 5)
	f.SetCellFormula("Data", "B1", "[.A1]*2")
	f.SetCellFormula("Sheet1", "A1", "[Data.B1]+10")

	err := f.RecalcAll()
	if err != nil {
		t.Fatalf("RecalcAll: %v", err)
	}

	v, _ := f.GetCellFloat("Sheet1", "A1")
	if v != 20 {
		t.Errorf("cross-sheet chain got %v, want 20", v)
	}
}
