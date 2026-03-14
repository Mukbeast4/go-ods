package goods

import (
	"errors"
	"testing"
)

func TestRangeSetStyle(t *testing.T) {
	f := NewFile()
	styleID, _ := f.NewStyle(&Style{Font: &Font{Bold: "true"}})

	err := f.Range("Sheet1", "A1:C3").SetStyle(styleID).Err()
	if err != nil {
		t.Fatalf("SetStyle() error = %v", err)
	}

	s := f.getSheet("Sheet1")
	for row := 1; row <= 3; row++ {
		for col := 1; col <= 3; col++ {
			c := s.getCell(col, row)
			if c == nil {
				t.Fatalf("cell (%d,%d) is nil", col, row)
			}
			if c.styleID != styleID {
				t.Errorf("cell (%d,%d) styleID = %d, want %d", col, row, c.styleID, styleID)
			}
		}
	}
}

func TestRangeSetValue(t *testing.T) {
	f := NewFile()

	err := f.Range("Sheet1", "B2:D4").SetValue(42).Err()
	if err != nil {
		t.Fatalf("SetValue() error = %v", err)
	}

	for row := 2; row <= 4; row++ {
		for col := 2; col <= 4; col++ {
			ref := Cell(col, row)
			val, _ := f.GetCellValue("Sheet1", ref)
			if val != "42" {
				t.Errorf("cell %s = %q, want %q", ref, val, "42")
			}
		}
	}
}

func TestRangeMerge(t *testing.T) {
	f := NewFile()

	err := f.Range("Sheet1", "A1:C3").Merge().Err()
	if err != nil {
		t.Fatalf("Merge() error = %v", err)
	}

	merges, _ := f.GetMergeCells("Sheet1")
	if len(merges) != 1 {
		t.Fatalf("expected 1 merge, got %d", len(merges))
	}
	if merges[0][0] != "A1" || merges[0][1] != "C3" {
		t.Errorf("merge = %v, want [A1 C3]", merges[0])
	}
}

func TestRangeMergeOverlap(t *testing.T) {
	f := NewFile()
	_ = f.Range("Sheet1", "A1:C3").Merge().Err()

	err := f.Range("Sheet1", "B2:D4").Merge().Err()
	if !errors.Is(err, ErrMergeOverlap) {
		t.Errorf("expected ErrMergeOverlap, got %v", err)
	}
}

func TestRangeSetNumberFormat(t *testing.T) {
	f := NewFile()
	_ = f.Range("Sheet1", "A1:B2").SetValue(1234.5).SetNumberFormat("#,##0.00").Err()

	format, _ := f.GetCellNumberFormat("Sheet1", "A1")
	if format != "#,##0.00" {
		t.Errorf("number format = %q, want %q", format, "#,##0.00")
	}
	format, _ = f.GetCellNumberFormat("Sheet1", "B2")
	if format != "#,##0.00" {
		t.Errorf("number format = %q, want %q", format, "#,##0.00")
	}
}

func TestRangeChaining(t *testing.T) {
	f := NewFile()
	styleID, _ := f.NewStyle(&Style{Fill: &Fill{Color: "#FF0000"}})

	err := f.Range("Sheet1", "A1:C1").
		SetValue("header").
		SetStyle(styleID).
		Merge().
		Err()
	if err != nil {
		t.Fatalf("chained operations error = %v", err)
	}

	val, _ := f.GetCellValue("Sheet1", "A1")
	if val != "header" {
		t.Errorf("value = %q, want %q", val, "header")
	}

	merges, _ := f.GetMergeCells("Sheet1")
	if len(merges) != 1 {
		t.Fatalf("expected 1 merge, got %d", len(merges))
	}
}

func TestRangeSheetNotFound(t *testing.T) {
	f := NewFile()

	err := f.Range("NonExistent", "A1:B2").SetValue(1).Err()
	if !errors.Is(err, ErrSheetNotFound) {
		t.Errorf("expected ErrSheetNotFound, got %v", err)
	}
}

func TestRangeFileClosed(t *testing.T) {
	f := NewFile()
	f.closed = true

	err := f.Range("Sheet1", "A1:B2").SetValue(1).Err()
	if !errors.Is(err, ErrFileClosed) {
		t.Errorf("expected ErrFileClosed, got %v", err)
	}
}

func TestRangeInvalidRef(t *testing.T) {
	f := NewFile()

	err := f.Range("Sheet1", "!!!").SetValue(1).Err()
	if err == nil {
		t.Error("expected error for invalid range ref")
	}
}

func TestRangeStyleNotFound(t *testing.T) {
	f := NewFile()

	err := f.Range("Sheet1", "A1:B2").SetStyle(999).Err()
	if !errors.Is(err, ErrStyleNotFound) {
		t.Errorf("expected ErrStyleNotFound, got %v", err)
	}
}

func TestRangeSingleCell(t *testing.T) {
	f := NewFile()

	err := f.Range("Sheet1", "B3").SetValue("single").Err()
	if err != nil {
		t.Fatalf("single cell range error = %v", err)
	}

	val, _ := f.GetCellValue("Sheet1", "B3")
	if val != "single" {
		t.Errorf("value = %q, want %q", val, "single")
	}
}
