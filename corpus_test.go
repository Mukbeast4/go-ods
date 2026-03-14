package goods

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
	"os"
	"testing"
)

func saveAndReopenCorpus(t *testing.T, f *File) *File {
	t.Helper()
	buf, err := f.WriteToBuffer()
	if err != nil {
		t.Fatalf("WriteToBuffer: %v", err)
	}
	f2, err := OpenBytes(buf.Bytes())
	if err != nil {
		t.Fatalf("OpenBytes: %v", err)
	}
	return f2
}

func extractContentXML(t *testing.T, data []byte) []byte {
	t.Helper()
	r, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("zip.NewReader: %v", err)
	}
	for _, f := range r.File {
		if f.Name == "content.xml" {
			rc, err := f.Open()
			if err != nil {
				t.Fatalf("open content.xml: %v", err)
			}
			defer rc.Close()
			b, err := io.ReadAll(rc)
			if err != nil {
				t.Fatalf("read content.xml: %v", err)
			}
			return b
		}
	}
	t.Fatal("content.xml not found in ODS")
	return nil
}

func TestCorpusRepeatedCells(t *testing.T) {
	f, err := OpenFile("testdata/corpus_repeated_cells.ods")
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	sheet := f.GetSheetList()[0]

	t.Run("repeated_columns", func(t *testing.T) {
		for _, col := range []string{"A", "B", "C", "D", "E"} {
			ref := col + "1"
			val, err := f.GetCellValue(sheet, ref)
			if err != nil {
				t.Fatalf("GetCellValue(%s): %v", ref, err)
			}
			if val != "42" {
				t.Errorf("%s = %q, want %q", ref, val, "42")
			}
		}
	})

	t.Run("end_marker", func(t *testing.T) {
		val, err := f.GetCellValue(sheet, "F1")
		if err != nil {
			t.Fatalf("GetCellValue(F1): %v", err)
		}
		if val != "end" {
			t.Errorf("F1 = %q, want %q", val, "end")
		}
	})

	t.Run("repeated_rows", func(t *testing.T) {
		for _, row := range []int{2, 3, 4} {
			ref := fmt.Sprintf("A%d", row)
			val, err := f.GetCellValue(sheet, ref)
			if err != nil {
				t.Fatalf("GetCellValue(%s): %v", ref, err)
			}
			if val != "repeated_row" {
				t.Errorf("%s = %q, want %q", ref, val, "repeated_row")
			}
		}
	})

	t.Run("after_repeated", func(t *testing.T) {
		val, err := f.GetCellValue(sheet, "A5")
		if err != nil {
			t.Fatalf("GetCellValue(A5): %v", err)
		}
		if val != "after_repeated" {
			t.Errorf("A5 = %q, want %q", val, "after_repeated")
		}
	})

	t.Run("round_trip", func(t *testing.T) {
		f2 := saveAndReopenCorpus(t, f)
		sheet2 := f2.GetSheetList()[0]

		for _, col := range []string{"A", "B", "C", "D", "E"} {
			ref := col + "1"
			val, err := f2.GetCellValue(sheet2, ref)
			if err != nil {
				t.Fatalf("GetCellValue(%s): %v", ref, err)
			}
			if val != "42" {
				t.Errorf("after round-trip %s = %q, want %q", ref, val, "42")
			}
		}

		val, err := f2.GetCellValue(sheet2, "F1")
		if err != nil {
			t.Fatalf("GetCellValue(F1): %v", err)
		}
		if val != "end" {
			t.Errorf("after round-trip F1 = %q, want %q", val, "end")
		}

		val, err = f2.GetCellValue(sheet2, "A5")
		if err != nil {
			t.Fatalf("GetCellValue(A5): %v", err)
		}
		if val != "after_repeated" {
			t.Errorf("after round-trip A5 = %q, want %q", val, "after_repeated")
		}
	})
}

func TestCorpusEmptyRows(t *testing.T) {
	f, err := OpenFile("testdata/corpus_empty_rows.ods")
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	sheet := f.GetSheetList()[0]

	t.Run("header", func(t *testing.T) {
		val, err := f.GetCellValue(sheet, "A1")
		if err != nil {
			t.Fatalf("GetCellValue(A1): %v", err)
		}
		if val != "header" {
			t.Errorf("A1 = %q, want %q", val, "header")
		}
	})

	t.Run("data_row", func(t *testing.T) {
		val, err := f.GetCellValue(sheet, "A2")
		if err != nil {
			t.Fatalf("GetCellValue(A2): %v", err)
		}
		if val != "100" {
			t.Errorf("A2 = %q, want %q", val, "100")
		}
	})

	t.Run("dimension", func(t *testing.T) {
		dim, err := f.GetSheetDimension(sheet)
		if err != nil {
			t.Fatalf("GetSheetDimension: %v", err)
		}
		t.Logf("sheet dimension: %s", dim)
	})
}

func TestCorpusNestedStyles(t *testing.T) {
	f, err := OpenFile("testdata/corpus_nested_styles.ods")
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	sheet := f.GetSheetList()[0]

	expected := map[string]string{
		"A1": "parent",
		"A2": "child",
		"A3": "grandchild",
	}

	t.Run("values", func(t *testing.T) {
		for ref, want := range expected {
			val, err := f.GetCellValue(sheet, ref)
			if err != nil {
				t.Fatalf("GetCellValue(%s): %v", ref, err)
			}
			if val != want {
				t.Errorf("%s = %q, want %q", ref, val, want)
			}
		}
	})

	t.Run("round_trip_values", func(t *testing.T) {
		f2 := saveAndReopenCorpus(t, f)
		sheet2 := f2.GetSheetList()[0]

		for ref, want := range expected {
			val, err := f2.GetCellValue(sheet2, ref)
			if err != nil {
				t.Fatalf("GetCellValue(%s): %v", ref, err)
			}
			if val != want {
				t.Errorf("after round-trip %s = %q, want %q", ref, val, want)
			}
		}
	})

	t.Run("style_references", func(t *testing.T) {
		buf, err := f.WriteToBuffer()
		if err != nil {
			t.Fatalf("WriteToBuffer: %v", err)
		}
		content := extractContentXML(t, buf.Bytes())

		for _, styleName := range []string{"ce-parent", "ce-child", "ce-grandchild"} {
			if !bytes.Contains(content, []byte(styleName)) {
				t.Errorf("content.xml does not contain style reference %q", styleName)
			}
		}
	})
}

func TestCorpusSpecialChars(t *testing.T) {
	f, err := OpenFile("testdata/corpus_special_chars.ods")
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	sheet := f.GetSheetList()[0]

	expected := map[string]string{
		"B1": "<less than>",
		"B2": "greater>than",
		"B3": "this & that",
		"B4": `say "hello"`,
		"B5": "it's here",
	}

	t.Run("values", func(t *testing.T) {
		for ref, want := range expected {
			val, err := f.GetCellValue(sheet, ref)
			if err != nil {
				t.Fatalf("GetCellValue(%s): %v", ref, err)
			}
			if val != want {
				t.Errorf("%s = %q, want %q", ref, val, want)
			}
		}
	})

	t.Run("round_trip", func(t *testing.T) {
		f2 := saveAndReopenCorpus(t, f)
		sheet2 := f2.GetSheetList()[0]

		for ref, want := range expected {
			val, err := f2.GetCellValue(sheet2, ref)
			if err != nil {
				t.Fatalf("GetCellValue(%s): %v", ref, err)
			}
			if val != want {
				t.Errorf("after round-trip %s = %q, want %q", ref, val, want)
			}
		}
	})
}

func TestCorpusSizeRatio(t *testing.T) {
	inputPath := "testdata/corpus_repeated_cells.ods"
	inputData, err := os.ReadFile(inputPath)
	if err != nil {
		t.Fatalf("ReadFile: %v", err)
	}
	inputSize := len(inputData)

	f, err := OpenFile(inputPath)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	buf, err := f.WriteToBuffer()
	if err != nil {
		t.Fatalf("WriteToBuffer: %v", err)
	}
	outputSize := buf.Len()

	ratio := float64(inputSize) / float64(outputSize)
	t.Logf("size ratio input/output: %.2f (input=%d, output=%d)", ratio, inputSize, outputSize)
}
