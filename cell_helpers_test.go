package goods

import (
	"path/filepath"
	"testing"
	"time"
)

func TestSetCellCurrencyDefault(t *testing.T) {
	f := NewFile()

	if err := f.SetCellCurrency("Sheet1", "A1", 1234.5, ""); err != nil {
		t.Fatalf("SetCellCurrency: %v", err)
	}

	got, err := f.GetCellFormattedValue("Sheet1", "A1")
	if err != nil {
		t.Fatalf("GetCellFormattedValue: %v", err)
	}
	if got != "$1,234.50" {
		t.Errorf("got %q, want %q", got, "$1,234.50")
	}

	ct, _ := f.GetCellType("Sheet1", "A1")
	if ct != CellTypeCurrency {
		t.Errorf("type = %v, want CellTypeCurrency", ct)
	}
}

func TestSetCellCurrencyEuro(t *testing.T) {
	f := NewFile()

	if err := f.SetCellCurrency("Sheet1", "B2", 999.99, "€"); err != nil {
		t.Fatalf("SetCellCurrency: %v", err)
	}

	got, _ := f.GetCellFormattedValue("Sheet1", "B2")
	if got != "€999.99" {
		t.Errorf("got %q, want %q", got, "€999.99")
	}

	fmtCode, _ := f.GetCellNumberFormat("Sheet1", "B2")
	if fmtCode != "€#,##0.00" {
		t.Errorf("format = %q, want %q", fmtCode, "€#,##0.00")
	}
}

func TestSetCellPercentageDecimals(t *testing.T) {
	f := NewFile()

	cases := []struct {
		cell     string
		value    float64
		decimals int
		want     string
	}{
		{"A1", 0.25, 0, "25%"},
		{"A2", 0.3333, 2, "33.33%"},
		{"A3", 1.0, 0, "100%"},
		{"A4", -0.05, 0, "-5%"},
	}

	for _, tc := range cases {
		if err := f.SetCellPercentage("Sheet1", tc.cell, tc.value, tc.decimals); err != nil {
			t.Fatalf("SetCellPercentage(%s): %v", tc.cell, err)
		}
		got, _ := f.GetCellFormattedValue("Sheet1", tc.cell)
		if got != tc.want {
			t.Errorf("cell %s: got %q, want %q", tc.cell, got, tc.want)
		}
		ct, _ := f.GetCellType("Sheet1", tc.cell)
		if ct != CellTypePercentage {
			t.Errorf("cell %s: type = %v, want CellTypePercentage", tc.cell, ct)
		}
	}
}

func TestSetCellDateTimeFormat(t *testing.T) {
	f := NewFile()

	tm := time.Date(2024, 3, 15, 9, 30, 0, 0, time.UTC)
	if err := f.SetCellDateTime("Sheet1", "A1", tm, "DD/MM/YYYY"); err != nil {
		t.Fatalf("SetCellDateTime: %v", err)
	}

	ct, _ := f.GetCellType("Sheet1", "A1")
	if ct != CellTypeDate {
		t.Errorf("type = %v, want CellTypeDate", ct)
	}

	fmtCode, _ := f.GetCellNumberFormat("Sheet1", "A1")
	if fmtCode != "DD/MM/YYYY" {
		t.Errorf("format = %q, want %q", fmtCode, "DD/MM/YYYY")
	}

	got, err := f.GetCellDate("Sheet1", "A1")
	if err != nil {
		t.Fatalf("GetCellDate: %v", err)
	}
	if !got.Equal(tm) {
		t.Errorf("got %v, want %v", got, tm)
	}
}

func TestSetCellDateTimeEmptyFormat(t *testing.T) {
	f := NewFile()

	tm := time.Date(2024, 1, 1, 0, 0, 0, 0, time.UTC)
	if err := f.SetCellDateTime("Sheet1", "A1", tm, ""); err != nil {
		t.Fatalf("SetCellDateTime: %v", err)
	}

	fmtCode, _ := f.GetCellNumberFormat("Sheet1", "A1")
	if fmtCode != "" {
		t.Errorf("format = %q, want empty", fmtCode)
	}
}

func TestSetCellDuration(t *testing.T) {
	f := NewFile()

	d := 1*time.Hour + 30*time.Minute + 45*time.Second
	if err := f.SetCellDuration("Sheet1", "A1", d); err != nil {
		t.Fatalf("SetCellDuration: %v", err)
	}

	ct, _ := f.GetCellType("Sheet1", "A1")
	if ct != CellTypeTime {
		t.Errorf("type = %v, want CellTypeTime", ct)
	}

	raw, _ := f.GetCellValue("Sheet1", "A1")
	if raw != "PT1H30M45S" {
		t.Errorf("raw = %q, want %q", raw, "PT1H30M45S")
	}

	fmtCode, _ := f.GetCellNumberFormat("Sheet1", "A1")
	if fmtCode != "[HH]:MM:SS" {
		t.Errorf("format = %q, want %q", fmtCode, "[HH]:MM:SS")
	}
}

func TestFormatISODurationNegative(t *testing.T) {
	if got := formatISODuration(-2 * time.Hour); got != "-PT2H0M0S" {
		t.Errorf("got %q, want %q", got, "-PT2H0M0S")
	}
}

func TestCellHelpersRoundTrip(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "helpers.ods")

	f := NewFile()
	if err := f.SetCellCurrency("Sheet1", "A1", 42.5, "$"); err != nil {
		t.Fatal(err)
	}
	if err := f.SetCellPercentage("Sheet1", "A2", 0.125, 2); err != nil {
		t.Fatal(err)
	}
	if err := f.SetCellDateTime("Sheet1", "A3", time.Date(2024, 6, 10, 0, 0, 0, 0, time.UTC), "DD/MM/YYYY"); err != nil {
		t.Fatal(err)
	}
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}

	reopened, err := OpenFile(path)
	if err != nil {
		t.Fatal(err)
	}

	if got, _ := reopened.GetCellFloat("Sheet1", "A1"); got != 42.5 {
		t.Errorf("A1 = %v, want 42.5", got)
	}
	if ct, _ := reopened.GetCellType("Sheet1", "A2"); ct != CellTypePercentage {
		t.Errorf("A2 type = %v, want CellTypePercentage", ct)
	}
	if ct, _ := reopened.GetCellType("Sheet1", "A3"); ct != CellTypeDate {
		t.Errorf("A3 type = %v, want CellTypeDate", ct)
	}
}

func TestCellHelpersErrorPaths(t *testing.T) {
	f := NewFile()

	if err := f.SetCellCurrency("Missing", "A1", 1, ""); err != ErrSheetNotFound {
		t.Errorf("SetCellCurrency on missing sheet: %v", err)
	}
	if err := f.SetCellPercentage("Missing", "A1", 0.1, 0); err != ErrSheetNotFound {
		t.Errorf("SetCellPercentage on missing sheet: %v", err)
	}
	if err := f.SetCellDateTime("Missing", "A1", time.Now(), ""); err != ErrSheetNotFound {
		t.Errorf("SetCellDateTime on missing sheet: %v", err)
	}
	if err := f.SetCellDuration("Missing", "A1", time.Second); err != ErrSheetNotFound {
		t.Errorf("SetCellDuration on missing sheet: %v", err)
	}

	if err := f.SetCellCurrency("Sheet1", "bogus", 1, ""); err == nil {
		t.Error("SetCellCurrency with bad cell should error")
	}
}

func TestCellHelpersFileClosed(t *testing.T) {
	f := NewFile()
	_ = f.Close()

	if err := f.SetCellCurrency("Sheet1", "A1", 1, ""); err != ErrFileClosed {
		t.Errorf("SetCellCurrency on closed file: %v", err)
	}
	if err := f.SetCellPercentage("Sheet1", "A1", 0.1, 0); err != ErrFileClosed {
		t.Errorf("SetCellPercentage on closed file: %v", err)
	}
	if err := f.SetCellDateTime("Sheet1", "A1", time.Now(), ""); err != ErrFileClosed {
		t.Errorf("SetCellDateTime on closed file: %v", err)
	}
	if err := f.SetCellDuration("Sheet1", "A1", time.Second); err != ErrFileClosed {
		t.Errorf("SetCellDuration on closed file: %v", err)
	}
}
