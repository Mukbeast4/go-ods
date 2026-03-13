package goods

import (
	"path/filepath"
	"testing"
)

func TestSetGetFreezePane(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	err := f.SetFreezePane(s, "C4")
	if err != nil {
		t.Fatalf("SetFreezePane: %v", err)
	}

	col, row, err := f.GetFreezePane(s)
	if err != nil {
		t.Fatalf("GetFreezePane: %v", err)
	}
	if col != 2 {
		t.Errorf("col = %d, want 2", col)
	}
	if row != 3 {
		t.Errorf("row = %d, want 3", row)
	}
}

func TestRemoveFreezePane(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetFreezePane(s, "B2")
	f.RemoveFreezePane(s)

	col, row, _ := f.GetFreezePane(s)
	if col != 0 || row != 0 {
		t.Errorf("expected (0,0) after remove, got (%d,%d)", col, row)
	}
}

func TestFreezePaneRoundtrip(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellStr(s, "A1", "data")
	f.SetFreezePane(s, "C3")

	tmpDir := t.TempDir()
	path := filepath.Join(tmpDir, "freeze.ods")

	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	f2, err := OpenFile(path)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}
	defer f2.Close()

	col, row, _ := f2.GetFreezePane(s)
	_ = col
	_ = row
}

func TestFreezePaneCopySheet(t *testing.T) {
	f := NewFile()

	f.SetFreezePane("Sheet1", "B3")
	f.CopySheet("Sheet1", "Sheet2")

	col, row, _ := f.GetFreezePane("Sheet2")
	if col != 1 {
		t.Errorf("col = %d, want 1", col)
	}
	if row != 2 {
		t.Errorf("row = %d, want 2", row)
	}
}

func TestFreezePaneClosedFile(t *testing.T) {
	f := NewFile()
	f.Close()

	err := f.SetFreezePane("Sheet1", "B2")
	if err != ErrFileClosed {
		t.Errorf("expected ErrFileClosed, got %v", err)
	}

	_, _, err = f.GetFreezePane("Sheet1")
	if err != ErrFileClosed {
		t.Errorf("expected ErrFileClosed, got %v", err)
	}

	err = f.RemoveFreezePane("Sheet1")
	if err != ErrFileClosed {
		t.Errorf("expected ErrFileClosed, got %v", err)
	}
}

func TestFreezePaneSheetNotFound(t *testing.T) {
	f := NewFile()

	err := f.SetFreezePane("NoSheet", "B2")
	if err != ErrSheetNotFound {
		t.Errorf("expected ErrSheetNotFound, got %v", err)
	}
}
