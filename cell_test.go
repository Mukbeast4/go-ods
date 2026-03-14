package goods

import (
	"errors"
	"testing"
)

func TestSetRowValues(t *testing.T) {
	f := NewFile()

	err := f.SetRowValues("Sheet1", 1, []any{"Name", "Age", 42, true})
	if err != nil {
		t.Fatalf("SetRowValues() error = %v", err)
	}

	val, _ := f.GetCellValue("Sheet1", "A1")
	if val != "Name" {
		t.Errorf("A1 = %q, want %q", val, "Name")
	}
	val, _ = f.GetCellValue("Sheet1", "B1")
	if val != "Age" {
		t.Errorf("B1 = %q, want %q", val, "Age")
	}
	val, _ = f.GetCellValue("Sheet1", "C1")
	if val != "42" {
		t.Errorf("C1 = %q, want %q", val, "42")
	}
	val, _ = f.GetCellValue("Sheet1", "D1")
	if val != "TRUE" {
		t.Errorf("D1 = %q, want %q", val, "TRUE")
	}
}

func TestSetRowValuesErrors(t *testing.T) {
	f := NewFile()

	err := f.SetRowValues("Sheet1", 0, []any{"a"})
	if !errors.Is(err, ErrRowOutOfRange) {
		t.Errorf("row=0: expected ErrRowOutOfRange, got %v", err)
	}

	err = f.SetRowValues("NoSheet", 1, []any{"a"})
	if !errors.Is(err, ErrSheetNotFound) {
		t.Errorf("bad sheet: expected ErrSheetNotFound, got %v", err)
	}

	f.closed = true
	err = f.SetRowValues("Sheet1", 1, []any{"a"})
	if !errors.Is(err, ErrFileClosed) {
		t.Errorf("closed: expected ErrFileClosed, got %v", err)
	}
}

func TestSetRowValuesCellError(t *testing.T) {
	f := NewFile()

	err := f.SetRowValues("NoSheet", 1, []any{"a"})
	var ce *CellError
	if !errors.As(err, &ce) {
		t.Fatalf("expected *CellError, got %T", err)
	}
	if ce.Sheet != "NoSheet" {
		t.Errorf("CellError.Sheet = %q, want %q", ce.Sheet, "NoSheet")
	}
}

func TestAppendRows(t *testing.T) {
	f := NewFile()

	_ = f.SetRowValues("Sheet1", 1, []any{"header1", "header2"})

	err := f.AppendRows("Sheet1", [][]any{
		{"a", 1},
		{"b", 2},
		{"c", 3},
	})
	if err != nil {
		t.Fatalf("AppendRows() error = %v", err)
	}

	val, _ := f.GetCellValue("Sheet1", "A2")
	if val != "a" {
		t.Errorf("A2 = %q, want %q", val, "a")
	}
	val, _ = f.GetCellValue("Sheet1", "B4")
	if val != "3" {
		t.Errorf("B4 = %q, want %q", val, "3")
	}
}

func TestAppendRowsEmptySheet(t *testing.T) {
	f := NewFile()

	err := f.AppendRows("Sheet1", [][]any{
		{"first", "row"},
	})
	if err != nil {
		t.Fatalf("AppendRows() error = %v", err)
	}

	val, _ := f.GetCellValue("Sheet1", "A1")
	if val != "first" {
		t.Errorf("A1 = %q, want %q", val, "first")
	}
}

func TestAppendRowsErrors(t *testing.T) {
	f := NewFile()

	err := f.AppendRows("NoSheet", [][]any{{"a"}})
	if !errors.Is(err, ErrSheetNotFound) {
		t.Errorf("bad sheet: expected ErrSheetNotFound, got %v", err)
	}

	f.closed = true
	err = f.AppendRows("Sheet1", [][]any{{"a"}})
	if !errors.Is(err, ErrFileClosed) {
		t.Errorf("closed: expected ErrFileClosed, got %v", err)
	}
}
