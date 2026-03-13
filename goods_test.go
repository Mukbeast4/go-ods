package goods

import (
	"os"
	"path/filepath"
	"testing"
	"time"
)

func TestNewFile(t *testing.T) {
	f := NewFile()
	if f == nil {
		t.Fatal("NewFile() returned nil")
	}

	sheets := f.GetSheetList()
	if len(sheets) != 1 || sheets[0] != "Sheet1" {
		t.Errorf("expected [Sheet1], got %v", sheets)
	}
}

func TestNewFileAndSave(t *testing.T) {
	f := NewFile()

	if err := f.SetCellValue("Sheet1", "A1", "Hello"); err != nil {
		t.Fatal(err)
	}
	if err := f.SetCellValue("Sheet1", "B1", 42); err != nil {
		t.Fatal(err)
	}
	if err := f.SetCellValue("Sheet1", "C1", 3.14); err != nil {
		t.Fatal(err)
	}
	if err := f.SetCellValue("Sheet1", "D1", true); err != nil {
		t.Fatal(err)
	}

	dir := t.TempDir()
	path := filepath.Join(dir, "test.ods")

	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}

	if _, err := os.Stat(path); os.IsNotExist(err) {
		t.Fatal("file not created")
	}
}

func TestFileRoundTrip(t *testing.T) {
	f := NewFile()

	f.SetCellValue("Sheet1", "A1", "Name")
	f.SetCellValue("Sheet1", "B1", "Age")
	f.SetCellValue("Sheet1", "C1", "Active")
	f.SetCellValue("Sheet1", "A2", "Alice")
	f.SetCellValue("Sheet1", "B2", 30)
	f.SetCellValue("Sheet1", "C2", true)
	f.SetCellValue("Sheet1", "A3", "Bob")
	f.SetCellValue("Sheet1", "B3", 25)
	f.SetCellValue("Sheet1", "C3", false)

	dir := t.TempDir()
	path := filepath.Join(dir, "roundtrip.ods")

	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}

	f2, err := OpenFile(path)
	if err != nil {
		t.Fatal(err)
	}

	val, err := f2.GetCellValue("Sheet1", "A1")
	if err != nil {
		t.Fatal(err)
	}
	if val != "Name" {
		t.Errorf("A1: expected 'Name', got %q", val)
	}

	val, err = f2.GetCellValue("Sheet1", "B2")
	if err != nil {
		t.Fatal(err)
	}
	if val != "30" {
		t.Errorf("B2: expected '30', got %q", val)
	}

	val, err = f2.GetCellValue("Sheet1", "C2")
	if err != nil {
		t.Fatal(err)
	}
	if val != "TRUE" {
		t.Errorf("C2: expected 'TRUE', got %q", val)
	}

	val, err = f2.GetCellValue("Sheet1", "C3")
	if err != nil {
		t.Fatal(err)
	}
	if val != "FALSE" {
		t.Errorf("C3: expected 'FALSE', got %q", val)
	}
}

func TestSheetOperations(t *testing.T) {
	f := NewFile()

	idx, err := f.NewSheet("Data")
	if err != nil {
		t.Fatal(err)
	}
	if idx != 1 {
		t.Errorf("expected index 1, got %d", idx)
	}

	sheets := f.GetSheetList()
	if len(sheets) != 2 {
		t.Fatalf("expected 2 sheets, got %d", len(sheets))
	}

	_, err = f.NewSheet("Data")
	if err != ErrSheetExists {
		t.Errorf("expected ErrSheetExists, got %v", err)
	}

	if err := f.SetSheetName("Data", "MyData"); err != nil {
		t.Fatal(err)
	}

	sheets = f.GetSheetList()
	found := false
	for _, s := range sheets {
		if s == "MyData" {
			found = true
		}
	}
	if !found {
		t.Error("renamed sheet not found")
	}

	if err := f.DeleteSheet("MyData"); err != nil {
		t.Fatal(err)
	}

	if f.SheetCount() != 1 {
		t.Errorf("expected 1 sheet, got %d", f.SheetCount())
	}

	err = f.DeleteSheet("Sheet1")
	if err != ErrNoSheets {
		t.Errorf("expected ErrNoSheets, got %v", err)
	}
}

func TestGetRows(t *testing.T) {
	f := NewFile()

	f.SetCellValue("Sheet1", "A1", "a")
	f.SetCellValue("Sheet1", "B1", "b")
	f.SetCellValue("Sheet1", "A2", "c")
	f.SetCellValue("Sheet1", "B2", "d")

	rows, err := f.GetRows("Sheet1")
	if err != nil {
		t.Fatal(err)
	}

	if len(rows) != 2 {
		t.Fatalf("expected 2 rows, got %d", len(rows))
	}
	if len(rows[0]) != 2 {
		t.Fatalf("expected 2 cols, got %d", len(rows[0]))
	}
	if rows[0][0] != "a" || rows[0][1] != "b" {
		t.Errorf("row 0: expected [a b], got %v", rows[0])
	}
	if rows[1][0] != "c" || rows[1][1] != "d" {
		t.Errorf("row 1: expected [c d], got %v", rows[1])
	}
}

func TestSetSheetRow(t *testing.T) {
	f := NewFile()

	err := f.SetSheetRow("Sheet1", "A1", []interface{}{"Name", "Age", true})
	if err != nil {
		t.Fatal(err)
	}

	val, _ := f.GetCellValue("Sheet1", "A1")
	if val != "Name" {
		t.Errorf("expected 'Name', got %q", val)
	}
	val, _ = f.GetCellValue("Sheet1", "C1")
	if val != "TRUE" {
		t.Errorf("expected 'TRUE', got %q", val)
	}
}

func TestTypedGetters(t *testing.T) {
	f := NewFile()

	f.SetCellFloat("Sheet1", "A1", 3.14)
	f.SetCellInt("Sheet1", "B1", 42)
	f.SetCellBool("Sheet1", "C1", true)
	f.SetCellDate("Sheet1", "D1", time.Date(2024, 1, 15, 0, 0, 0, 0, time.UTC))

	fv, err := f.GetCellFloat("Sheet1", "A1")
	if err != nil {
		t.Fatal(err)
	}
	if fv != 3.14 {
		t.Errorf("expected 3.14, got %f", fv)
	}

	iv, err := f.GetCellInt("Sheet1", "B1")
	if err != nil {
		t.Fatal(err)
	}
	if iv != 42 {
		t.Errorf("expected 42, got %d", iv)
	}

	bv, err := f.GetCellBool("Sheet1", "C1")
	if err != nil {
		t.Fatal(err)
	}
	if !bv {
		t.Error("expected true, got false")
	}

	dv, err := f.GetCellDate("Sheet1", "D1")
	if err != nil {
		t.Fatal(err)
	}
	if dv.Year() != 2024 || dv.Month() != 1 || dv.Day() != 15 {
		t.Errorf("expected 2024-01-15, got %v", dv)
	}
}

func TestMergeCells(t *testing.T) {
	f := NewFile()

	f.SetCellValue("Sheet1", "A1", "Merged")

	if err := f.MergeCell("Sheet1", "A1", "C3"); err != nil {
		t.Fatal(err)
	}

	merges, err := f.GetMergeCells("Sheet1")
	if err != nil {
		t.Fatal(err)
	}
	if len(merges) != 1 {
		t.Fatalf("expected 1 merge, got %d", len(merges))
	}
	if merges[0][0] != "A1" || merges[0][1] != "C3" {
		t.Errorf("expected [A1 C3], got %v", merges[0])
	}

	err = f.MergeCell("Sheet1", "B2", "D4")
	if err != ErrMergeOverlap {
		t.Errorf("expected ErrMergeOverlap, got %v", err)
	}

	if err := f.UnmergeCell("Sheet1", "A1", "C3"); err != nil {
		t.Fatal(err)
	}

	merges, _ = f.GetMergeCells("Sheet1")
	if len(merges) != 0 {
		t.Errorf("expected 0 merges, got %d", len(merges))
	}
}

func TestFormula(t *testing.T) {
	f := NewFile()

	f.SetCellValue("Sheet1", "A1", 10)
	f.SetCellValue("Sheet1", "A2", 20)
	f.SetCellFormula("Sheet1", "A3", "SUM(A1:A2)")

	formula, err := f.GetCellFormula("Sheet1", "A3")
	if err != nil {
		t.Fatal(err)
	}
	if formula != "SUM(A1:A2)" {
		t.Errorf("expected 'SUM(A1:A2)', got %q", formula)
	}

	dir := t.TempDir()
	path := filepath.Join(dir, "formula.ods")

	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}

	f2, err := OpenFile(path)
	if err != nil {
		t.Fatal(err)
	}

	formula, err = f2.GetCellFormula("Sheet1", "A3")
	if err != nil {
		t.Fatal(err)
	}
	if formula != "SUM(A1:A2)" {
		t.Errorf("round-trip formula: expected 'SUM(A1:A2)', got %q", formula)
	}
}

func TestRowIterator(t *testing.T) {
	f := NewFile()

	f.SetCellValue("Sheet1", "A1", "a")
	f.SetCellValue("Sheet1", "A3", "c")

	it, err := f.NewRowIterator("Sheet1")
	if err != nil {
		t.Fatal(err)
	}

	var rowIndices []int
	for it.Next() {
		rowIndices = append(rowIndices, it.RowIndex())
	}

	if len(rowIndices) != 2 {
		t.Fatalf("expected 2 rows, got %d", len(rowIndices))
	}
	if rowIndices[0] != 1 || rowIndices[1] != 3 {
		t.Errorf("expected [1, 3], got %v", rowIndices)
	}
}

func TestDocProperties(t *testing.T) {
	f := NewFile()

	f.SetDocProperties(&DocProperties{
		Title:   "Test Doc",
		Creator: "goods",
	})

	props, err := f.GetDocProperties()
	if err != nil {
		t.Fatal(err)
	}
	if props.Title != "Test Doc" {
		t.Errorf("expected 'Test Doc', got %q", props.Title)
	}
	if props.Creator != "goods" {
		t.Errorf("expected 'goods', got %q", props.Creator)
	}
}

func TestStyleAndApply(t *testing.T) {
	f := NewFile()

	styleID, err := f.NewStyle(&Style{
		Font: &Font{
			Family: "Arial",
			Size:   "12pt",
			Bold:   "bold",
			Color:  "#FF0000",
		},
		Fill: &Fill{
			Color: "#FFFF00",
		},
	})
	if err != nil {
		t.Fatal(err)
	}

	f.SetCellValue("Sheet1", "A1", "Styled")

	if err := f.SetCellStyle("Sheet1", "A1", "A1", styleID); err != nil {
		t.Fatal(err)
	}

	dir := t.TempDir()
	path := filepath.Join(dir, "styled.ods")
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}

	if _, err := os.Stat(path); os.IsNotExist(err) {
		t.Fatal("styled file not created")
	}
}

func TestWriteToBuffer(t *testing.T) {
	f := NewFile()
	f.SetCellValue("Sheet1", "A1", "buffered")

	buf, err := f.WriteToBuffer()
	if err != nil {
		t.Fatal(err)
	}

	if buf.Len() == 0 {
		t.Error("buffer is empty")
	}

	f2, err := OpenBytes(buf.Bytes())
	if err != nil {
		t.Fatal(err)
	}

	val, _ := f2.GetCellValue("Sheet1", "A1")
	if val != "buffered" {
		t.Errorf("expected 'buffered', got %q", val)
	}
}

func TestCopySheet(t *testing.T) {
	f := NewFile()
	f.SetCellValue("Sheet1", "A1", "original")

	if err := f.CopySheet("Sheet1", "Sheet1 Copy"); err != nil {
		t.Fatal(err)
	}

	val, err := f.GetCellValue("Sheet1 Copy", "A1")
	if err != nil {
		t.Fatal(err)
	}
	if val != "original" {
		t.Errorf("expected 'original', got %q", val)
	}
}

func TestClosedFile(t *testing.T) {
	f := NewFile()
	f.Close()

	if err := f.SetCellValue("Sheet1", "A1", "x"); err != ErrFileClosed {
		t.Errorf("expected ErrFileClosed, got %v", err)
	}
}

func TestInsertRows(t *testing.T) {
	f := NewFile()
	f.SetCellValue("Sheet1", "A1", "row1")
	f.SetCellValue("Sheet1", "A2", "row2")
	f.SetCellValue("Sheet1", "A3", "row3")

	if err := f.InsertRows("Sheet1", 2, 2); err != nil {
		t.Fatal(err)
	}

	val, _ := f.GetCellValue("Sheet1", "A1")
	if val != "row1" {
		t.Errorf("A1: expected 'row1', got %q", val)
	}

	val, _ = f.GetCellValue("Sheet1", "A4")
	if val != "row2" {
		t.Errorf("A4: expected 'row2', got %q", val)
	}

	val, _ = f.GetCellValue("Sheet1", "A5")
	if val != "row3" {
		t.Errorf("A5: expected 'row3', got %q", val)
	}
}

func TestRemoveRow(t *testing.T) {
	f := NewFile()
	f.SetCellValue("Sheet1", "A1", "row1")
	f.SetCellValue("Sheet1", "A2", "row2")
	f.SetCellValue("Sheet1", "A3", "row3")

	if err := f.RemoveRow("Sheet1", 2); err != nil {
		t.Fatal(err)
	}

	val, _ := f.GetCellValue("Sheet1", "A2")
	if val != "row3" {
		t.Errorf("A2 after remove: expected 'row3', got %q", val)
	}
}

func TestSetColWidth(t *testing.T) {
	f := NewFile()

	if err := f.SetColWidth("Sheet1", "A", 5.0); err != nil {
		t.Fatal(err)
	}

	width, err := f.GetColWidth("Sheet1", "A")
	if err != nil {
		t.Fatal(err)
	}
	if width != 5.0 {
		t.Errorf("expected 5.0, got %f", width)
	}
}
