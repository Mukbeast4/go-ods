package goods

import (
	"path/filepath"
	"testing"
)

func TestSetColVisible(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	if err := f.SetColVisible(s, "B", false); err != nil {
		t.Fatalf("SetColVisible: %v", err)
	}

	visible, err := f.GetColVisible(s, "B")
	if err != nil {
		t.Fatalf("GetColVisible: %v", err)
	}
	if visible {
		t.Error("expected column B to be hidden")
	}
}

func TestGetColVisibleDefault(t *testing.T) {
	f := NewFile()

	visible, err := f.GetColVisible("Sheet1", "A")
	if err != nil {
		t.Fatalf("GetColVisible: %v", err)
	}
	if !visible {
		t.Error("expected default column to be visible")
	}
}

func TestSetRowVisible(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	if err := f.SetRowVisible(s, 3, false); err != nil {
		t.Fatalf("SetRowVisible: %v", err)
	}

	visible, err := f.GetRowVisible(s, 3)
	if err != nil {
		t.Fatalf("GetRowVisible: %v", err)
	}
	if visible {
		t.Error("expected row 3 to be hidden")
	}
}

func TestGetRowVisibleDefault(t *testing.T) {
	f := NewFile()

	visible, err := f.GetRowVisible("Sheet1", 5)
	if err != nil {
		t.Fatalf("GetRowVisible: %v", err)
	}
	if !visible {
		t.Error("expected default row to be visible")
	}
}

func TestColVisibilityRoundTrip(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellStr(s, "A1", "a")
	f.SetCellStr(s, "B1", "b")
	f.SetCellStr(s, "C1", "c")
	f.SetColVisible(s, "B", false)

	tmpDir := t.TempDir()
	path := filepath.Join(tmpDir, "col_vis.ods")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	f2, err := OpenFile(path)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}
	defer f2.Close()

	visible, err := f2.GetColVisible(s, "B")
	if err != nil {
		t.Fatalf("GetColVisible: %v", err)
	}
	if visible {
		t.Error("expected column B to remain hidden after round-trip")
	}

	visibleA, _ := f2.GetColVisible(s, "A")
	if !visibleA {
		t.Error("expected column A to remain visible after round-trip")
	}
}

func TestRowVisibilityRoundTrip(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellStr(s, "A1", "row1")
	f.SetCellStr(s, "A2", "row2")
	f.SetCellStr(s, "A3", "row3")
	f.SetRowVisible(s, 2, false)

	tmpDir := t.TempDir()
	path := filepath.Join(tmpDir, "row_vis.ods")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	f2, err := OpenFile(path)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}
	defer f2.Close()

	visible, err := f2.GetRowVisible(s, 2)
	if err != nil {
		t.Fatalf("GetRowVisible: %v", err)
	}
	if visible {
		t.Error("expected row 2 to remain hidden after round-trip")
	}

	visible1, _ := f2.GetRowVisible(s, 1)
	if !visible1 {
		t.Error("expected row 1 to remain visible after round-trip")
	}
}

func TestHiddenRowGetRows(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellStr(s, "A1", "visible")
	f.SetCellStr(s, "A2", "hidden")
	f.SetCellStr(s, "A3", "visible")
	f.SetRowVisible(s, 2, false)

	rows, err := f.GetRows(s)
	if err != nil {
		t.Fatalf("GetRows: %v", err)
	}
	if len(rows) < 3 {
		t.Fatalf("expected at least 3 rows, got %d", len(rows))
	}
	if rows[1][0] != "hidden" {
		t.Errorf("hidden row data = %q, want %q", rows[1][0], "hidden")
	}
}

func TestSetColAutoFit(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	if err := f.SetColAutoFit(s, "A", true); err != nil {
		t.Fatalf("SetColAutoFit: %v", err)
	}

	autoFit, err := f.GetColAutoFit(s, "A")
	if err != nil {
		t.Fatalf("GetColAutoFit: %v", err)
	}
	if !autoFit {
		t.Error("expected column A to be auto-fit")
	}
}

func TestColAutoFitRoundTrip(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellStr(s, "A1", "test")
	f.SetColAutoFit(s, "A", true)

	tmpDir := t.TempDir()
	path := filepath.Join(tmpDir, "autofit.ods")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	f2, err := OpenFile(path)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}
	defer f2.Close()

	autoFit, err := f2.GetColAutoFit(s, "A")
	if err != nil {
		t.Fatalf("GetColAutoFit: %v", err)
	}
	if !autoFit {
		t.Error("expected column A to remain auto-fit after round-trip")
	}
}

func TestColAutoFitWithWidth(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetColWidth(s, "A", 5.0)
	f.SetColAutoFit(s, "A", true)

	w, _ := f.GetColWidth(s, "A")
	if w != 5.0 {
		t.Errorf("width = %f, want 5.0", w)
	}

	af, _ := f.GetColAutoFit(s, "A")
	if !af {
		t.Error("expected auto-fit to be true")
	}
}
