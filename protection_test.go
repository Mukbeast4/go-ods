package goods

import (
	"path/filepath"
	"testing"
)

func TestSetSheetProtection(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	if err := f.SetSheetProtection(s, true); err != nil {
		t.Fatalf("SetSheetProtection: %v", err)
	}

	protected, err := f.IsSheetProtected(s)
	if err != nil {
		t.Fatalf("IsSheetProtected: %v", err)
	}
	if !protected {
		t.Error("expected sheet to be protected")
	}
}

func TestSheetProtectionRoundTrip(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellStr(s, "A1", "data")
	f.SetSheetProtection(s, true)

	tmpDir := t.TempDir()
	path := filepath.Join(tmpDir, "protected.ods")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	f2, err := OpenFile(path)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}
	defer f2.Close()

	protected, err := f2.IsSheetProtected(s)
	if err != nil {
		t.Fatalf("IsSheetProtected: %v", err)
	}
	if !protected {
		t.Error("expected sheet to remain protected after round-trip")
	}
}

func TestCellProtectionStyle(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	protectedVal := true
	styleID, err := f.NewStyle(&Style{
		Protected: &protectedVal,
	})
	if err != nil {
		t.Fatalf("NewStyle: %v", err)
	}

	f.SetCellStr(s, "A1", "locked")
	f.SetCellStyle(s, "A1", "A1", styleID)

	tmpDir := t.TempDir()
	path := filepath.Join(tmpDir, "cell_protect.ods")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	f2, err := OpenFile(path)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}
	defer f2.Close()

	val, _ := f2.GetCellValue(s, "A1")
	if val != "locked" {
		t.Errorf("cell value = %q, want %q", val, "locked")
	}
}

func TestProtectedSheetDefaultCells(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellStr(s, "A1", "data")
	f.SetSheetProtection(s, true)

	val, _ := f.GetCellValue(s, "A1")
	if val != "data" {
		t.Errorf("cell value = %q, want %q", val, "data")
	}
}

func TestSheetProtectionClosedFile(t *testing.T) {
	f := NewFile()
	f.Close()

	err := f.SetSheetProtection("Sheet1", true)
	if err != ErrFileClosed {
		t.Errorf("expected ErrFileClosed, got %v", err)
	}

	_, err = f.IsSheetProtected("Sheet1")
	if err != ErrFileClosed {
		t.Errorf("expected ErrFileClosed, got %v", err)
	}
}
