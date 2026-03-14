package goods

import (
	"path/filepath"
	"testing"
)

func TestSetConditionalFormat(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	rules := []ConditionalRule{
		{Value: "of:cell-content()>90", StyleName: "good", BaseCellAddress: "Sheet1.A1"},
	}

	if err := f.SetConditionalFormat(s, "A1:A100", rules); err != nil {
		t.Fatalf("SetConditionalFormat: %v", err)
	}

	formats, err := f.GetConditionalFormats(s)
	if err != nil {
		t.Fatalf("GetConditionalFormats: %v", err)
	}
	if len(formats) != 1 {
		t.Fatalf("expected 1 format, got %d", len(formats))
	}
	if formats[0].Range != "A1:A100" {
		t.Errorf("range = %q, want %q", formats[0].Range, "A1:A100")
	}
	if len(formats[0].Rules) != 1 {
		t.Fatalf("expected 1 rule, got %d", len(formats[0].Rules))
	}
}

func TestConditionalFormatRoundTrip(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellStr(s, "A1", "data")
	rules := []ConditionalRule{
		{Value: "of:cell-content()>90", StyleName: "good", BaseCellAddress: "Sheet1.A1"},
	}
	f.SetConditionalFormat(s, "A1:A100", rules)

	tmpDir := t.TempDir()
	path := filepath.Join(tmpDir, "conditional.ods")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	f2, err := OpenFile(path)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}
	defer f2.Close()

	formats, err := f2.GetConditionalFormats(s)
	if err != nil {
		t.Fatalf("GetConditionalFormats: %v", err)
	}
	if len(formats) != 1 {
		t.Fatalf("expected 1 format, got %d", len(formats))
	}
	if len(formats[0].Rules) != 1 {
		t.Fatalf("expected 1 rule, got %d", len(formats[0].Rules))
	}
	if formats[0].Rules[0].StyleName != "good" {
		t.Errorf("style name = %q, want %q", formats[0].Rules[0].StyleName, "good")
	}
}

func TestMultipleConditionalRules(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	rules := []ConditionalRule{
		{Value: "of:cell-content()>90", StyleName: "good", BaseCellAddress: "Sheet1.A1"},
		{Value: "of:cell-content()<50", StyleName: "bad", BaseCellAddress: "Sheet1.A1"},
	}
	f.SetConditionalFormat(s, "A1:A100", rules)

	formats, _ := f.GetConditionalFormats(s)
	if len(formats[0].Rules) != 2 {
		t.Fatalf("expected 2 rules, got %d", len(formats[0].Rules))
	}
}

func TestRemoveConditionalFormat(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	rules := []ConditionalRule{
		{Value: "of:cell-content()>90", StyleName: "good", BaseCellAddress: "Sheet1.A1"},
	}
	f.SetConditionalFormat(s, "A1:A100", rules)

	err := f.RemoveConditionalFormat(s, "A1:A100")
	if err != nil {
		t.Fatalf("RemoveConditionalFormat: %v", err)
	}

	formats, _ := f.GetConditionalFormats(s)
	if len(formats) != 0 {
		t.Errorf("expected 0 formats after remove, got %d", len(formats))
	}
}

func TestRemoveConditionalFormatNotFound(t *testing.T) {
	f := NewFile()

	err := f.RemoveConditionalFormat("Sheet1", "A1:A100")
	if err != ErrConditionalFormatNotFound {
		t.Errorf("expected ErrConditionalFormatNotFound, got %v", err)
	}
}

func TestConditionalFormatWithStyle(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellStr(s, "A1", "data")
	rules := []ConditionalRule{
		{Value: "of:cell-content()>90", StyleName: "highlight", BaseCellAddress: "Sheet1.A1"},
	}
	f.SetConditionalFormat(s, "A1:A10", rules)

	tmpDir := t.TempDir()
	path := filepath.Join(tmpDir, "cond_style.ods")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	f2, err := OpenFile(path)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}
	defer f2.Close()

	formats, _ := f2.GetConditionalFormats(s)
	if len(formats) != 1 {
		t.Fatalf("expected 1 format, got %d", len(formats))
	}
	if formats[0].Rules[0].StyleName != "highlight" {
		t.Errorf("style = %q, want %q", formats[0].Rules[0].StyleName, "highlight")
	}
}
