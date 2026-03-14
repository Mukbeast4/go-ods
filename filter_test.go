package goods

import (
	"path/filepath"
	"testing"
)

func TestSetFilterCriteria(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetAutoFilter(s, "A1", "C10")

	criteria := []FilterCriteria{
		{Column: 0, Values: []string{"Alice"}},
	}
	if err := f.SetFilterCriteria(s, criteria); err != nil {
		t.Fatalf("SetFilterCriteria: %v", err)
	}

	got, err := f.GetFilterCriteria(s)
	if err != nil {
		t.Fatalf("GetFilterCriteria: %v", err)
	}
	if len(got) != 1 {
		t.Fatalf("expected 1 criteria, got %d", len(got))
	}
	if got[0].Column != 0 {
		t.Errorf("column = %d, want 0", got[0].Column)
	}
	if len(got[0].Values) != 1 || got[0].Values[0] != "Alice" {
		t.Errorf("values = %v, want [Alice]", got[0].Values)
	}
}

func TestFilterCriteriaRoundTrip(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellStr(s, "A1", "Name")
	f.SetCellStr(s, "A2", "Alice")
	f.SetAutoFilter(s, "A1", "A10")
	f.SetFilterCriteria(s, []FilterCriteria{
		{Column: 0, Values: []string{"Alice"}},
	})

	tmpDir := t.TempDir()
	path := filepath.Join(tmpDir, "filter.ods")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	f2, err := OpenFile(path)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}
	defer f2.Close()

	got, err := f2.GetFilterCriteria(s)
	if err != nil {
		t.Fatalf("GetFilterCriteria: %v", err)
	}
	if len(got) != 1 {
		t.Fatalf("expected 1 criteria, got %d", len(got))
	}
	if got[0].Values[0] != "Alice" {
		t.Errorf("value = %q, want %q", got[0].Values[0], "Alice")
	}
}

func TestClearFilterCriteria(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetAutoFilter(s, "A1", "C10")
	f.SetFilterCriteria(s, []FilterCriteria{
		{Column: 0, Values: []string{"Alice"}},
	})

	if err := f.ClearFilterCriteria(s); err != nil {
		t.Fatalf("ClearFilterCriteria: %v", err)
	}

	got, _ := f.GetFilterCriteria(s)
	if len(got) != 0 {
		t.Errorf("expected 0 criteria after clear, got %d", len(got))
	}
}

func TestSetSort(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	keys := []SortKey{
		{Column: 1, Descending: true},
	}
	if err := f.SetSort(s, keys); err != nil {
		t.Fatalf("SetSort: %v", err)
	}

	got, err := f.GetSort(s)
	if err != nil {
		t.Fatalf("GetSort: %v", err)
	}
	if len(got) != 1 {
		t.Fatalf("expected 1 sort key, got %d", len(got))
	}
	if got[0].Column != 1 {
		t.Errorf("column = %d, want 1", got[0].Column)
	}
	if !got[0].Descending {
		t.Error("expected descending to be true")
	}
}

func TestSortRoundTrip(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellStr(s, "A1", "Name")
	f.SetCellStr(s, "B1", "Score")
	f.SetAutoFilter(s, "A1", "B10")
	f.SetSort(s, []SortKey{
		{Column: 1, Descending: true},
	})

	tmpDir := t.TempDir()
	path := filepath.Join(tmpDir, "sort.ods")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	f2, err := OpenFile(path)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}
	defer f2.Close()

	got, err := f2.GetSort(s)
	if err != nil {
		t.Fatalf("GetSort: %v", err)
	}
	if len(got) != 1 {
		t.Fatalf("expected 1 sort key, got %d", len(got))
	}
	if !got[0].Descending {
		t.Error("expected descending after round-trip")
	}
}

func TestSortAndFilter(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellStr(s, "A1", "Name")
	f.SetAutoFilter(s, "A1", "B10")
	f.SetFilterCriteria(s, []FilterCriteria{
		{Column: 0, Values: []string{"Alice"}},
	})
	f.SetSort(s, []SortKey{
		{Column: 1, Descending: false},
	})

	tmpDir := t.TempDir()
	path := filepath.Join(tmpDir, "sort_filter.ods")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	f2, err := OpenFile(path)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}
	defer f2.Close()

	criteria, _ := f2.GetFilterCriteria(s)
	if len(criteria) != 1 {
		t.Fatalf("expected 1 filter criteria, got %d", len(criteria))
	}

	sortKeys, _ := f2.GetSort(s)
	if len(sortKeys) != 1 {
		t.Fatalf("expected 1 sort key, got %d", len(sortKeys))
	}
}

func TestRemoveSort(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetSort(s, []SortKey{{Column: 0}})
	f.RemoveSort(s)

	got, _ := f.GetSort(s)
	if len(got) != 0 {
		t.Errorf("expected 0 sort keys after remove, got %d", len(got))
	}
}

func TestFilterCriteriaNoAutoFilter(t *testing.T) {
	f := NewFile()

	err := f.SetFilterCriteria("Sheet1", []FilterCriteria{
		{Column: 0, Values: []string{"test"}},
	})
	if err != ErrAutoFilterNotFound {
		t.Errorf("expected ErrAutoFilterNotFound, got %v", err)
	}
}
