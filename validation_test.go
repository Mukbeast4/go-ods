package goods

import (
	"path/filepath"
	"testing"
)

func TestSetGetValidation(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	err := f.SetDataValidation(s, "A1", "A10", &DataValidation{
		Type:         "whole-number",
		Operator:     "between",
		Formula1:     "1",
		Formula2:     "100",
		AllowEmpty:   true,
		ErrorTitle:   "Error",
		ErrorMessage: "Value must be between 1 and 100",
		ErrorStyle:   "stop",
	})
	if err != nil {
		t.Fatalf("SetDataValidation: %v", err)
	}

	dv, err := f.GetDataValidation(s, "A1")
	if err != nil {
		t.Fatalf("GetDataValidation: %v", err)
	}
	if dv == nil {
		t.Fatal("expected validation, got nil")
	}
	if dv.Type != "whole-number" {
		t.Errorf("Type = %q, want whole-number", dv.Type)
	}
	if dv.Operator != "between" {
		t.Errorf("Operator = %q, want between", dv.Operator)
	}
	if dv.Formula1 != "1" {
		t.Errorf("Formula1 = %q, want 1", dv.Formula1)
	}
	if dv.Formula2 != "100" {
		t.Errorf("Formula2 = %q, want 100", dv.Formula2)
	}
}

func TestRemoveValidation(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetDataValidation(s, "A1", "A5", &DataValidation{
		Type:     "list",
		Formula1: "\"Yes\",\"No\"",
	})

	err := f.RemoveDataValidation(s, "A1", "A5")
	if err != nil {
		t.Fatalf("RemoveDataValidation: %v", err)
	}

	dv, _ := f.GetDataValidation(s, "A1")
	if dv != nil {
		t.Error("expected nil after remove")
	}
}

func TestValidationNoValidation(t *testing.T) {
	f := NewFile()

	dv, err := f.GetDataValidation("Sheet1", "A1")
	if err != nil {
		t.Fatalf("GetDataValidation: %v", err)
	}
	if dv != nil {
		t.Error("expected nil for cell without validation")
	}
}

func TestValidationRoundtrip(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellStr(s, "A1", "data")
	f.SetDataValidation(s, "A1", "A5", &DataValidation{
		Type:         "whole-number",
		Operator:     "between",
		Formula1:     "1",
		Formula2:     "100",
		AllowEmpty:   true,
		ErrorTitle:   "Error",
		ErrorMessage: "Out of range",
		ErrorStyle:   "stop",
	})

	tmpDir := t.TempDir()
	path := filepath.Join(tmpDir, "validation.ods")

	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	f2, err := OpenFile(path)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}
	defer f2.Close()

	dv, err := f2.GetDataValidation(s, "A1")
	if err != nil {
		t.Fatalf("GetDataValidation: %v", err)
	}
	if dv == nil {
		t.Fatal("expected validation after roundtrip")
	}
	if dv.Type != "whole-number" {
		t.Errorf("Type = %q, want whole-number", dv.Type)
	}
	if dv.Operator != "between" {
		t.Errorf("Operator = %q, want between", dv.Operator)
	}
}

func TestValidationClosedFile(t *testing.T) {
	f := NewFile()
	f.Close()

	err := f.SetDataValidation("Sheet1", "A1", "A5", &DataValidation{Type: "list"})
	if err != ErrFileClosed {
		t.Errorf("expected ErrFileClosed, got %v", err)
	}
}
