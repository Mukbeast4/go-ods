package goods

import (
	"path/filepath"
	"testing"
)

func TestSetGetAutoFilter(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	err := f.SetAutoFilter(s, "A1", "D100")
	if err != nil {
		t.Fatalf("SetAutoFilter: %v", err)
	}

	tl, br, err := f.GetAutoFilter(s)
	if err != nil {
		t.Fatalf("GetAutoFilter: %v", err)
	}
	if tl != "A1" {
		t.Errorf("topLeft = %q, want A1", tl)
	}
	if br != "D100" {
		t.Errorf("bottomRight = %q, want D100", br)
	}
}

func TestRemoveAutoFilter(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetAutoFilter(s, "A1", "D10")
	err := f.RemoveAutoFilter(s)
	if err != nil {
		t.Fatalf("RemoveAutoFilter: %v", err)
	}

	_, _, err = f.GetAutoFilter(s)
	if err != ErrAutoFilterNotFound {
		t.Errorf("expected ErrAutoFilterNotFound, got %v", err)
	}
}

func TestAutoFilterNotFound(t *testing.T) {
	f := NewFile()

	_, _, err := f.GetAutoFilter("Sheet1")
	if err != ErrAutoFilterNotFound {
		t.Errorf("expected ErrAutoFilterNotFound, got %v", err)
	}
}

func TestAutoFilterUpdate(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetAutoFilter(s, "A1", "D10")
	f.SetAutoFilter(s, "A1", "F20")

	_, br, _ := f.GetAutoFilter(s)
	if br != "F20" {
		t.Errorf("bottomRight = %q, want F20 after update", br)
	}
}

func TestAutoFilterRoundtrip(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellStr(s, "A1", "header")
	f.SetAutoFilter(s, "A1", "D10")

	tmpDir := t.TempDir()
	path := filepath.Join(tmpDir, "autofilter.ods")

	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	f2, err := OpenFile(path)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}
	defer f2.Close()

	tl, br, err := f2.GetAutoFilter(s)
	if err != nil {
		t.Fatalf("GetAutoFilter: %v", err)
	}
	if tl != "A1" {
		t.Errorf("topLeft = %q, want A1", tl)
	}
	if br != "D10" {
		t.Errorf("bottomRight = %q, want D10", br)
	}
}

func TestAutoFilterDeleteSheet(t *testing.T) {
	f := NewFile()
	f.NewSheet("Sheet2")

	f.SetAutoFilter("Sheet1", "A1", "D10")
	f.SetAutoFilter("Sheet2", "A1", "C5")

	f.DeleteSheet("Sheet1")

	_, _, err := f.GetAutoFilter("Sheet2")
	if err != nil {
		t.Fatalf("expected auto-filter on Sheet2, got %v", err)
	}
}

func TestAutoFilterClosedFile(t *testing.T) {
	f := NewFile()
	f.Close()

	err := f.SetAutoFilter("Sheet1", "A1", "D10")
	if err != ErrFileClosed {
		t.Errorf("expected ErrFileClosed, got %v", err)
	}

	_, _, err = f.GetAutoFilter("Sheet1")
	if err != ErrFileClosed {
		t.Errorf("expected ErrFileClosed, got %v", err)
	}

	err = f.RemoveAutoFilter("Sheet1")
	if err != ErrFileClosed {
		t.Errorf("expected ErrFileClosed, got %v", err)
	}
}

func TestAutoFilterInsertRows(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetAutoFilter(s, "A2", "D10")
	f.InsertRows(s, 2, 3)

	tl, br, _ := f.GetAutoFilter(s)
	if tl != "A5" {
		t.Errorf("topLeft = %q, want A5 after InsertRows", tl)
	}
	if br != "D13" {
		t.Errorf("bottomRight = %q, want D13 after InsertRows", br)
	}
}
