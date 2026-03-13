package goods

import "testing"

func TestShiftFormulaInsertRow(t *testing.T) {
	result := shiftFormulaRefs("[.A5]", "Sheet1", true, 3, 3)
	if result != "[.A8]" {
		t.Errorf("got %q, want [.A8]", result)
	}
}

func TestShiftFormulaDeleteRow(t *testing.T) {
	result := shiftFormulaRefs("[.A5]", "Sheet1", true, 3, -1)
	if result != "[.A4]" {
		t.Errorf("got %q, want [.A4]", result)
	}
}

func TestShiftFormulaInsertCol(t *testing.T) {
	result := shiftFormulaRefs("[.C1]", "Sheet1", false, 2, 1)
	if result != "[.D1]" {
		t.Errorf("got %q, want [.D1]", result)
	}
}

func TestShiftFormulaDeleteCol(t *testing.T) {
	result := shiftFormulaRefs("[.C1]", "Sheet1", false, 2, -1)
	if result != "[.B1]" {
		t.Errorf("got %q, want [.B1]", result)
	}
}

func TestShiftFormulaAbsoluteRow(t *testing.T) {
	result := shiftFormulaRefs("[.$A$5]", "Sheet1", true, 3, 3)
	if result != "[.$A$5]" {
		t.Errorf("got %q, want [.$A$5]", result)
	}
}

func TestShiftFormulaAbsoluteCol(t *testing.T) {
	result := shiftFormulaRefs("[.$C1]", "Sheet1", false, 2, 1)
	if result != "[.$C1]" {
		t.Errorf("got %q, want [.$C1]", result)
	}
}

func TestShiftFormulaRange(t *testing.T) {
	result := shiftFormulaRefs("[.A1:.B5]", "Sheet1", true, 3, 3)
	if result != "[.A1:.B8]" {
		t.Errorf("got %q, want [.A1:.B8]", result)
	}
}

func TestShiftFormulaCrossSheet(t *testing.T) {
	result := shiftFormulaRefs("[Sheet2.A5]", "Sheet1", true, 3, 3)
	if result != "[Sheet2.A5]" {
		t.Errorf("got %q, want [Sheet2.A5]", result)
	}
}

func TestShiftFormulaDeleteRef(t *testing.T) {
	result := shiftFormulaRefs("[.A3]", "Sheet1", true, 3, -1)
	if result != "#REF!" {
		t.Errorf("got %q, want #REF!", result)
	}
}

func TestShiftFormulaString(t *testing.T) {
	result := shiftFormulaRefs("\"texte [.A1]\"", "Sheet1", true, 1, 1)
	if result != "\"texte [.A1]\"" {
		t.Errorf("got %q, want \"texte [.A1]\"", result)
	}
}

func TestShiftFormulaNoShift(t *testing.T) {
	result := shiftFormulaRefs("[.A1]", "Sheet1", true, 5, 2)
	if result != "[.A1]" {
		t.Errorf("got %q, want [.A1]", result)
	}
}

func TestShiftFormulaCrossSheetSameSheet(t *testing.T) {
	result := shiftFormulaRefs("[Sheet2.A5]", "Sheet2", true, 3, 3)
	if result != "[Sheet2.A8]" {
		t.Errorf("got %q, want [Sheet2.A8]", result)
	}
}

func TestShiftFormulaComplex(t *testing.T) {
	result := shiftFormulaRefs("SUM([.A1:.A10])+[.B5]*2", "Sheet1", true, 3, 2)
	if result != "SUM([.A1:.A12])+[.B7]*2" {
		t.Errorf("got %q, want SUM([.A1:.A12])+[.B7]*2", result)
	}
}

func TestInsertRowsShiftsFormulas(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellFloat(s, "A1", 10)
	f.SetCellFloat(s, "A2", 20)
	f.SetCellFloat(s, "A3", 30)
	f.SetCellFormula(s, "B1", "[.A1]+[.A2]+[.A3]")

	f.InsertRows(s, 2, 2)

	formula, _ := f.GetCellFormula(s, "B1")
	expected := "[.A1]+[.A4]+[.A5]"
	if formula != expected {
		t.Errorf("InsertRows formula got %q, want %q", formula, expected)
	}
}

func TestRemoveRowShiftsFormulas(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellFloat(s, "A1", 10)
	f.SetCellFloat(s, "A2", 20)
	f.SetCellFloat(s, "A3", 30)
	f.SetCellFormula(s, "B1", "[.A1]+[.A3]")

	f.RemoveRow(s, 2)

	formula, _ := f.GetCellFormula(s, "B1")
	expected := "[.A1]+[.A2]"
	if formula != expected {
		t.Errorf("RemoveRow formula got %q, want %q", formula, expected)
	}
}

func TestInsertColsShiftsFormulas(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellFloat(s, "A1", 10)
	f.SetCellFloat(s, "B1", 20)
	f.SetCellFloat(s, "C1", 30)
	f.SetCellFormula(s, "D1", "[.A1]+[.B1]+[.C1]")

	f.InsertCols(s, "B", 1)

	formula, _ := f.GetCellFormula(s, "E1")
	expected := "[.A1]+[.C1]+[.D1]"
	if formula != expected {
		t.Errorf("InsertCols formula got %q, want %q", formula, expected)
	}
}

func TestShiftFormulaDeleteRange(t *testing.T) {
	result := shiftFormulaRefs("[.A3:.B3]", "Sheet1", true, 3, -1)
	if result != "#REF!" {
		t.Errorf("got %q, want #REF!", result)
	}
}
