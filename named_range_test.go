package goods

import (
	"path/filepath"
	"testing"
)

func TestSetGetNamedRange(t *testing.T) {
	f := NewFile()

	err := f.SetNamedRange("MyRange", "Sheet1", "A1", "D10")
	if err != nil {
		t.Fatalf("SetNamedRange: %v", err)
	}

	nr, err := f.GetNamedRange("MyRange")
	if err != nil {
		t.Fatalf("GetNamedRange: %v", err)
	}
	if nr.Name != "MyRange" {
		t.Errorf("Name = %q, want MyRange", nr.Name)
	}
	if nr.Sheet != "Sheet1" {
		t.Errorf("Sheet = %q, want Sheet1", nr.Sheet)
	}
	if nr.TopLeft != "A1" {
		t.Errorf("TopLeft = %q, want A1", nr.TopLeft)
	}
	if nr.BottomRight != "D10" {
		t.Errorf("BottomRight = %q, want D10", nr.BottomRight)
	}
}

func TestDeleteNamedRange(t *testing.T) {
	f := NewFile()

	f.SetNamedRange("MyRange", "Sheet1", "A1", "D10")
	err := f.DeleteNamedRange("MyRange")
	if err != nil {
		t.Fatalf("DeleteNamedRange: %v", err)
	}

	_, err = f.GetNamedRange("MyRange")
	if err != ErrNamedRangeNotFound {
		t.Errorf("expected ErrNamedRangeNotFound, got %v", err)
	}
}

func TestDeleteNamedRangeNotFound(t *testing.T) {
	f := NewFile()

	err := f.DeleteNamedRange("NoRange")
	if err != ErrNamedRangeNotFound {
		t.Errorf("expected ErrNamedRangeNotFound, got %v", err)
	}
}

func TestGetNamedRanges(t *testing.T) {
	f := NewFile()

	f.SetNamedRange("Range1", "Sheet1", "A1", "B5")
	f.SetNamedRange("Range2", "Sheet1", "C1", "D5")

	ranges := f.GetNamedRanges()
	if len(ranges) != 2 {
		t.Fatalf("expected 2 ranges, got %d", len(ranges))
	}
}

func TestNamedRangeUpdate(t *testing.T) {
	f := NewFile()

	f.SetNamedRange("MyRange", "Sheet1", "A1", "D10")
	f.SetNamedRange("MyRange", "Sheet1", "A1", "F20")

	nr, _ := f.GetNamedRange("MyRange")
	if nr.BottomRight != "F20" {
		t.Errorf("BottomRight = %q, want F20 after update", nr.BottomRight)
	}

	ranges := f.GetNamedRanges()
	if len(ranges) != 1 {
		t.Errorf("expected 1 range after update, got %d", len(ranges))
	}
}

func TestNamedRangeRoundtrip(t *testing.T) {
	f := NewFile()

	f.SetCellStr("Sheet1", "A1", "data")
	f.SetNamedRange("TestRange", "Sheet1", "A1", "C5")

	tmpDir := t.TempDir()
	path := filepath.Join(tmpDir, "named_range.ods")

	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	f2, err := OpenFile(path)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}
	defer f2.Close()

	nr, err := f2.GetNamedRange("TestRange")
	if err != nil {
		t.Fatalf("GetNamedRange: %v", err)
	}
	if nr.Sheet != "Sheet1" {
		t.Errorf("Sheet = %q, want Sheet1", nr.Sheet)
	}
	if nr.TopLeft != "A1" {
		t.Errorf("TopLeft = %q, want A1", nr.TopLeft)
	}
	if nr.BottomRight != "C5" {
		t.Errorf("BottomRight = %q, want C5", nr.BottomRight)
	}
}

func TestNamedRangeInsertRows(t *testing.T) {
	f := NewFile()

	f.SetNamedRange("MyRange", "Sheet1", "B2", "D5")
	f.InsertRows("Sheet1", 2, 3)

	nr, _ := f.GetNamedRange("MyRange")
	if nr.TopLeft != "B5" {
		t.Errorf("TopLeft = %q, want B5 after InsertRows", nr.TopLeft)
	}
	if nr.BottomRight != "D8" {
		t.Errorf("BottomRight = %q, want D8 after InsertRows", nr.BottomRight)
	}
}

func TestNamedRangeDeleteSheet(t *testing.T) {
	f := NewFile()
	f.NewSheet("Sheet2")

	f.SetNamedRange("Range1", "Sheet1", "A1", "B2")
	f.SetNamedRange("Range2", "Sheet2", "A1", "B2")

	f.DeleteSheet("Sheet1")

	ranges := f.GetNamedRanges()
	if len(ranges) != 1 {
		t.Fatalf("expected 1 range after delete sheet, got %d", len(ranges))
	}
	if ranges[0].Name != "Range2" {
		t.Errorf("expected Range2, got %s", ranges[0].Name)
	}
}

func TestNamedRangeClosedFile(t *testing.T) {
	f := NewFile()
	f.Close()

	err := f.SetNamedRange("R", "Sheet1", "A1", "B2")
	if err != ErrFileClosed {
		t.Errorf("expected ErrFileClosed, got %v", err)
	}

	_, err = f.GetNamedRange("R")
	if err != ErrFileClosed {
		t.Errorf("expected ErrFileClosed, got %v", err)
	}

	err = f.DeleteNamedRange("R")
	if err != ErrFileClosed {
		t.Errorf("expected ErrFileClosed, got %v", err)
	}
}
