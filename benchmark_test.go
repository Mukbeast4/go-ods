package goods

import (
	"fmt"
	"os"
	"path/filepath"
	"testing"
)

func BenchmarkOpenFile(b *testing.B) {
	for _, tc := range []struct {
		name, path string
	}{
		{"small", "testdata/sample.ods"},
		{"medium", "testdata/golden_formulas.ods"},
		{"large", "testdata/golden_large.ods"},
	} {
		data, err := os.ReadFile(tc.path)
		if err != nil {
			b.Fatalf("read %s: %v", tc.path, err)
		}
		b.Run(tc.name, func(b *testing.B) {
			b.ReportAllocs()
			for i := 0; i < b.N; i++ {
				f, err := OpenBytes(data)
				if err != nil {
					b.Fatal(err)
				}
				_ = f
			}
		})
	}
}

func BenchmarkWriteToBuffer(b *testing.B) {
	for _, tc := range []struct {
		name, path string
	}{
		{"small", "testdata/sample.ods"},
		{"medium", "testdata/golden_formulas.ods"},
		{"large", "testdata/golden_large.ods"},
	} {
		data, err := os.ReadFile(tc.path)
		if err != nil {
			b.Fatalf("read %s: %v", tc.path, err)
		}
		f, err := OpenBytes(data)
		if err != nil {
			b.Fatalf("open %s: %v", tc.path, err)
		}
		b.Run(tc.name, func(b *testing.B) {
			b.ReportAllocs()
			for i := 0; i < b.N; i++ {
				buf, err := f.WriteToBuffer()
				if err != nil {
					b.Fatal(err)
				}
				_ = buf
			}
		})
	}
}

func BenchmarkSaveAs(b *testing.B) {
	data, err := os.ReadFile("testdata/golden_large.ods")
	if err != nil {
		b.Fatalf("read: %v", err)
	}
	f, err := OpenBytes(data)
	if err != nil {
		b.Fatal(err)
	}
	tmpDir, err := os.MkdirTemp("", "bench-saveas-*")
	if err != nil {
		b.Fatal(err)
	}
	defer os.RemoveAll(tmpDir)

	b.ReportAllocs()
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		path := filepath.Join(tmpDir, fmt.Sprintf("out_%d.ods", i))
		if err := f.SaveAs(path); err != nil {
			b.Fatal(err)
		}
	}
}

func BenchmarkRecalcAll(b *testing.B) {
	data, err := os.ReadFile("testdata/golden_formulas.ods")
	if err != nil {
		b.Fatalf("read: %v", err)
	}

	b.ReportAllocs()
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		f, err := OpenBytes(data)
		if err != nil {
			b.Fatal(err)
		}
		if err := f.RecalcAll(); err != nil {
			b.Fatal(err)
		}
	}
}

func BenchmarkRecalcLarge(b *testing.B) {
	f := NewFile()
	for row := 1; row <= 1000; row++ {
		cellA, err := CoordinatesToCellName(1, row)
		if err != nil {
			b.Fatal(err)
		}
		if err := f.SetCellFloat("Sheet1", cellA, float64(row)*1.5); err != nil {
			b.Fatal(err)
		}

		cellB, err := CoordinatesToCellName(2, row)
		if err != nil {
			b.Fatal(err)
		}
		formula := fmt.Sprintf("SUM([.A1:.A%d])", row)
		if err := f.SetCellFormula("Sheet1", cellB, formula); err != nil {
			b.Fatal(err)
		}
	}

	b.ReportAllocs()
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		if err := f.RecalcAll(); err != nil {
			b.Fatal(err)
		}
	}
}

func BenchmarkRoundTrip(b *testing.B) {
	data, err := os.ReadFile("testdata/golden_large.ods")
	if err != nil {
		b.Fatalf("read: %v", err)
	}

	b.ReportAllocs()
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		f, err := OpenBytes(data)
		if err != nil {
			b.Fatal(err)
		}
		if err := f.RecalcAll(); err != nil {
			b.Fatal(err)
		}
		buf, err := f.WriteToBuffer()
		if err != nil {
			b.Fatal(err)
		}
		_ = buf
	}
}

func BenchmarkSetCellValue(b *testing.B) {
	b.ReportAllocs()
	for i := 0; i < b.N; i++ {
		b.StopTimer()
		f := NewFile()
		b.StartTimer()

		for j := 1; j <= 10000; j++ {
			cellRef, err := CoordinatesToCellName(1, j)
			if err != nil {
				b.Fatal(err)
			}
			if err := f.SetCellValue("Sheet1", cellRef, j); err != nil {
				b.Fatal(err)
			}
		}
	}
}

func BenchmarkGetRows(b *testing.B) {
	data, err := os.ReadFile("testdata/golden_large.ods")
	if err != nil {
		b.Fatalf("read: %v", err)
	}
	f, err := OpenBytes(data)
	if err != nil {
		b.Fatal(err)
	}

	b.ReportAllocs()
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		rows, err := f.GetRows("Sheet1")
		if err != nil {
			b.Fatal(err)
		}
		_ = rows
	}
}
