package goods

import (
	"path/filepath"
	"testing"
)

func TestSetGetNumberFormat(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellFloat(s, "A1", 1234.56)
	f.SetCellNumberFormat(s, "A1", "#,##0.00")

	fmt, err := f.GetCellNumberFormat(s, "A1")
	if err != nil {
		t.Fatalf("GetCellNumberFormat: %v", err)
	}
	if fmt != "#,##0.00" {
		t.Errorf("format = %q, want #,##0.00", fmt)
	}
}

func TestFormattedValueThousands(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellFloat(s, "A1", 1234567.89)
	f.SetCellNumberFormat(s, "A1", "#,##0.00")

	v, err := f.GetCellFormattedValue(s, "A1")
	if err != nil {
		t.Fatalf("GetCellFormattedValue: %v", err)
	}
	if v != "1,234,567.89" {
		t.Errorf("formatted = %q, want 1,234,567.89", v)
	}
}

func TestFormattedValuePercentage(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellFloat(s, "A1", 0.75)
	f.SetCellNumberFormat(s, "A1", "0%")

	v, _ := f.GetCellFormattedValue(s, "A1")
	if v != "75%" {
		t.Errorf("formatted = %q, want 75%%", v)
	}
}

func TestFormattedValuePercentageDecimals(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellFloat(s, "A1", 0.1234)
	f.SetCellNumberFormat(s, "A1", "0.00%")

	v, _ := f.GetCellFormattedValue(s, "A1")
	if v != "12.34%" {
		t.Errorf("formatted = %q, want 12.34%%", v)
	}
}

func TestFormattedValueCurrency(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellFloat(s, "A1", 1234.56)
	f.SetCellNumberFormat(s, "A1", "$#,##0.00")

	v, _ := f.GetCellFormattedValue(s, "A1")
	if v != "$1,234.56" {
		t.Errorf("formatted = %q, want $1,234.56", v)
	}
}

func TestFormattedValueEuro(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellFloat(s, "A1", 1234.56)
	f.SetCellNumberFormat(s, "A1", "€#,##0.00")

	v, _ := f.GetCellFormattedValue(s, "A1")
	if v != "€1,234.56" {
		t.Errorf("formatted = %q, want €1,234.56", v)
	}
}

func TestFormattedValueNoFormat(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellFloat(s, "A1", 42)

	v, _ := f.GetCellFormattedValue(s, "A1")
	if v != "42" {
		t.Errorf("formatted = %q, want 42", v)
	}
}

func TestFormattedValueRoundTrip(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellFloat(s, "A1", 1234.56)
	f.SetCellNumberFormat(s, "A1", "#,##0.00")

	tmpDir := t.TempDir()
	path := filepath.Join(tmpDir, "format.ods")

	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	f2, err := OpenFile(path)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}
	defer f2.Close()

	v, _ := f2.GetCellFloat(s, "A1")
	if v != 1234.56 {
		t.Errorf("value after roundtrip = %v, want 1234.56", v)
	}
}

func TestNumberFormatClosedFile(t *testing.T) {
	f := NewFile()
	f.Close()

	err := f.SetCellNumberFormat("Sheet1", "A1", "0.00")
	if err != ErrFileClosed {
		t.Errorf("expected ErrFileClosed, got %v", err)
	}

	_, err = f.GetCellNumberFormat("Sheet1", "A1")
	if err != ErrFileClosed {
		t.Errorf("expected ErrFileClosed, got %v", err)
	}

	_, err = f.GetCellFormattedValue("Sheet1", "A1")
	if err != ErrFileClosed {
		t.Errorf("expected ErrFileClosed, got %v", err)
	}
}

func TestNumberFormatSheetNotFound(t *testing.T) {
	f := NewFile()

	err := f.SetCellNumberFormat("NoSheet", "A1", "0.00")
	if err != ErrSheetNotFound {
		t.Errorf("expected ErrSheetNotFound, got %v", err)
	}
}

func TestFormatWithThousandsInteger(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellFloat(s, "A1", 1000000)
	f.SetCellNumberFormat(s, "A1", "#,##0")

	v, _ := f.GetCellFormattedValue(s, "A1")
	if v != "1,000,000" {
		t.Errorf("formatted = %q, want 1,000,000", v)
	}
}

func TestFormatFixedDecimals(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellFloat(s, "A1", 3.1)
	f.SetCellNumberFormat(s, "A1", "0.00")

	v, _ := f.GetCellFormattedValue(s, "A1")
	if v != "3.10" {
		t.Errorf("formatted = %q, want 3.10", v)
	}
}
