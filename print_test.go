package goods

import (
	"path/filepath"
	"testing"
)

func TestSetGetPrintRange(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	err := f.SetPrintRange(s, "A1", "D10")
	if err != nil {
		t.Fatalf("SetPrintRange: %v", err)
	}

	tl, br, err := f.GetPrintRange(s)
	if err != nil {
		t.Fatalf("GetPrintRange: %v", err)
	}
	if tl != "A1" {
		t.Errorf("topLeft = %q, want A1", tl)
	}
	if br != "D10" {
		t.Errorf("bottomRight = %q, want D10", br)
	}
}

func TestRemovePrintRange(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetPrintRange(s, "A1", "D10")
	f.RemovePrintRange(s)

	tl, br, _ := f.GetPrintRange(s)
	if tl != "" || br != "" {
		t.Errorf("expected empty after remove, got %q:%q", tl, br)
	}
}

func TestPrintRangeNone(t *testing.T) {
	f := NewFile()

	tl, br, err := f.GetPrintRange("Sheet1")
	if err != nil {
		t.Fatalf("GetPrintRange: %v", err)
	}
	if tl != "" || br != "" {
		t.Errorf("expected empty, got %q:%q", tl, br)
	}
}

func TestPrintRangeRoundtrip(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellStr(s, "A1", "data")
	f.SetPrintRange(s, "A1", "C5")

	tmpDir := t.TempDir()
	path := filepath.Join(tmpDir, "print.ods")

	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	f2, err := OpenFile(path)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}
	defer f2.Close()

	tl, br, err := f2.GetPrintRange(s)
	if err != nil {
		t.Fatalf("GetPrintRange: %v", err)
	}
	if tl != "A1" {
		t.Errorf("topLeft = %q, want A1", tl)
	}
	if br != "C5" {
		t.Errorf("bottomRight = %q, want C5", br)
	}
}

func TestPageSetup(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	err := f.SetPageSetup(s, &PageSetup{
		Orientation:  "landscape",
		PaperWidth:   "29.7cm",
		PaperHeight:  "21.001cm",
		MarginTop:    1.0,
		MarginBottom: 1.0,
		MarginLeft:   1.5,
		MarginRight:  1.5,
	})
	if err != nil {
		t.Fatalf("SetPageSetup: %v", err)
	}

	ps, err := f.GetPageSetup(s)
	if err != nil {
		t.Fatalf("GetPageSetup: %v", err)
	}
	if ps == nil {
		t.Fatal("expected page setup, got nil")
	}
	if ps.Orientation != "landscape" {
		t.Errorf("Orientation = %q, want landscape", ps.Orientation)
	}
	if ps.MarginLeft != 1.5 {
		t.Errorf("MarginLeft = %f, want 1.5", ps.MarginLeft)
	}
}

func TestPageSetupNone(t *testing.T) {
	f := NewFile()

	ps, err := f.GetPageSetup("Sheet1")
	if err != nil {
		t.Fatalf("GetPageSetup: %v", err)
	}
	if ps != nil {
		t.Error("expected nil for sheet without page setup")
	}
}

func TestPrintRangeInsertRows(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetPrintRange(s, "A2", "D10")
	f.InsertRows(s, 2, 3)

	tl, br, _ := f.GetPrintRange(s)
	if tl != "A5" {
		t.Errorf("topLeft = %q, want A5 after InsertRows", tl)
	}
	if br != "D13" {
		t.Errorf("bottomRight = %q, want D13 after InsertRows", br)
	}
}

func TestPrintRangeCopySheet(t *testing.T) {
	f := NewFile()

	f.SetPrintRange("Sheet1", "A1", "D10")
	f.CopySheet("Sheet1", "Sheet2")

	tl, br, _ := f.GetPrintRange("Sheet2")
	if tl != "A1" {
		t.Errorf("topLeft = %q, want A1", tl)
	}
	if br != "D10" {
		t.Errorf("bottomRight = %q, want D10", br)
	}
}

func TestPrintRangeClosedFile(t *testing.T) {
	f := NewFile()
	f.Close()

	err := f.SetPrintRange("Sheet1", "A1", "D10")
	if err != ErrFileClosed {
		t.Errorf("expected ErrFileClosed, got %v", err)
	}

	_, _, err = f.GetPrintRange("Sheet1")
	if err != ErrFileClosed {
		t.Errorf("expected ErrFileClosed, got %v", err)
	}

	err = f.RemovePrintRange("Sheet1")
	if err != ErrFileClosed {
		t.Errorf("expected ErrFileClosed, got %v", err)
	}
}
