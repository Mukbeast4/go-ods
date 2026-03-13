package goods

import (
	"bytes"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func TestExportCSVBasic(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellStr(s, "A1", "Name")
	f.SetCellStr(s, "B1", "Age")
	f.SetCellStr(s, "A2", "Alice")
	f.SetCellFloat(s, "B2", 30)
	f.SetCellStr(s, "A3", "Bob")
	f.SetCellFloat(s, "B3", 25)

	var buf bytes.Buffer
	err := f.ExportCSV(s, &buf, nil)
	if err != nil {
		t.Fatalf("ExportCSV: %v", err)
	}

	output := buf.String()
	lines := strings.Split(strings.TrimSpace(output), "\n")
	if len(lines) != 3 {
		t.Fatalf("expected 3 lines, got %d: %q", len(lines), output)
	}
	if lines[0] != "Name,Age" {
		t.Errorf("line 0 = %q, want Name,Age", lines[0])
	}
	if lines[1] != "Alice,30" {
		t.Errorf("line 1 = %q, want Alice,30", lines[1])
	}
}

func TestExportCSVCustomSeparator(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellStr(s, "A1", "a")
	f.SetCellStr(s, "B1", "b")

	var buf bytes.Buffer
	err := f.ExportCSV(s, &buf, &CSVOptions{Separator: ';'})
	if err != nil {
		t.Fatalf("ExportCSV: %v", err)
	}

	if strings.TrimSpace(buf.String()) != "a;b" {
		t.Errorf("got %q, want a;b", buf.String())
	}
}

func TestExportCSVFile(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellStr(s, "A1", "hello")
	f.SetCellStr(s, "B1", "world")

	tmpDir := t.TempDir()
	path := filepath.Join(tmpDir, "test.csv")

	err := f.ExportCSVFile(s, path, nil)
	if err != nil {
		t.Fatalf("ExportCSVFile: %v", err)
	}

	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatalf("ReadFile: %v", err)
	}

	if strings.TrimSpace(string(data)) != "hello,world" {
		t.Errorf("file content = %q, want hello,world", string(data))
	}
}

func TestExportCSVEmpty(t *testing.T) {
	f := NewFile()

	var buf bytes.Buffer
	err := f.ExportCSV("Sheet1", &buf, nil)
	if err != nil {
		t.Fatalf("ExportCSV: %v", err)
	}

	if buf.Len() != 0 {
		t.Errorf("expected empty output for empty sheet, got %q", buf.String())
	}
}

func TestExportCSVSheetNotFound(t *testing.T) {
	f := NewFile()

	var buf bytes.Buffer
	err := f.ExportCSV("NoSheet", &buf, nil)
	if err != ErrSheetNotFound {
		t.Errorf("expected ErrSheetNotFound, got %v", err)
	}
}

func TestExportCSVClosedFile(t *testing.T) {
	f := NewFile()
	f.Close()

	var buf bytes.Buffer
	err := f.ExportCSV("Sheet1", &buf, nil)
	if err != ErrFileClosed {
		t.Errorf("expected ErrFileClosed, got %v", err)
	}
}

func TestExportCSVCRLF(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellStr(s, "A1", "a")
	f.SetCellStr(s, "A2", "b")

	var buf bytes.Buffer
	err := f.ExportCSV(s, &buf, &CSVOptions{UseCRLF: true})
	if err != nil {
		t.Fatalf("ExportCSV: %v", err)
	}

	if !strings.Contains(buf.String(), "\r\n") {
		t.Errorf("expected CRLF line endings")
	}
}
