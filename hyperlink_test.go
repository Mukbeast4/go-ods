package goods

import (
	"path/filepath"
	"testing"
)

func TestSetGetHyperlink(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	err := f.SetCellHyperlink(s, "A1", "https://example.com", "Example")
	if err != nil {
		t.Fatalf("SetCellHyperlink: %v", err)
	}

	url, display, err := f.GetCellHyperlink(s, "A1")
	if err != nil {
		t.Fatalf("GetCellHyperlink: %v", err)
	}
	if url != "https://example.com" {
		t.Errorf("URL = %q, want https://example.com", url)
	}
	if display != "Example" {
		t.Errorf("Display = %q, want Example", display)
	}
}

func TestHyperlinkDefaultDisplay(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellHyperlink(s, "A1", "https://example.com", "")

	url, display, _ := f.GetCellHyperlink(s, "A1")
	if url != "https://example.com" {
		t.Errorf("URL = %q, want https://example.com", url)
	}
	if display != "https://example.com" {
		t.Errorf("Display = %q, want https://example.com", display)
	}
}

func TestRemoveHyperlink(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellHyperlink(s, "A1", "https://example.com", "Example")
	f.RemoveCellHyperlink(s, "A1")

	url, _, _ := f.GetCellHyperlink(s, "A1")
	if url != "" {
		t.Errorf("URL should be empty after remove, got %q", url)
	}
}

func TestHyperlinkNoLink(t *testing.T) {
	f := NewFile()

	url, display, err := f.GetCellHyperlink("Sheet1", "A1")
	if err != nil {
		t.Fatalf("GetCellHyperlink: %v", err)
	}
	if url != "" || display != "" {
		t.Errorf("expected empty, got URL=%q display=%q", url, display)
	}
}

func TestHyperlinkRoundtrip(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellHyperlink(s, "A1", "https://example.com", "Example")
	f.SetCellStr(s, "B1", "normal")

	tmpDir := t.TempDir()
	path := filepath.Join(tmpDir, "hyperlink.ods")

	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	f2, err := OpenFile(path)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}
	defer f2.Close()

	url, display, err := f2.GetCellHyperlink(s, "A1")
	if err != nil {
		t.Fatalf("GetCellHyperlink: %v", err)
	}
	if url != "https://example.com" {
		t.Errorf("URL after roundtrip = %q, want https://example.com", url)
	}
	if display != "Example" {
		t.Errorf("Display after roundtrip = %q, want Example", display)
	}
}

func TestHyperlinkClosedFile(t *testing.T) {
	f := NewFile()
	f.Close()

	err := f.SetCellHyperlink("Sheet1", "A1", "https://example.com", "")
	if err != ErrFileClosed {
		t.Errorf("expected ErrFileClosed, got %v", err)
	}

	_, _, err = f.GetCellHyperlink("Sheet1", "A1")
	if err != ErrFileClosed {
		t.Errorf("expected ErrFileClosed, got %v", err)
	}

	err = f.RemoveCellHyperlink("Sheet1", "A1")
	if err != ErrFileClosed {
		t.Errorf("expected ErrFileClosed, got %v", err)
	}
}

func TestHyperlinkSheetNotFound(t *testing.T) {
	f := NewFile()

	err := f.SetCellHyperlink("NoSheet", "A1", "https://example.com", "")
	if err != ErrSheetNotFound {
		t.Errorf("expected ErrSheetNotFound, got %v", err)
	}
}
