package goods

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func FuzzOpenBytes(f *testing.F) {
	entries, err := os.ReadDir("testdata")
	if err != nil {
		f.Fatalf("read testdata: %v", err)
	}
	for _, e := range entries {
		if strings.HasSuffix(e.Name(), ".ods") {
			data, err := os.ReadFile(filepath.Join("testdata", e.Name()))
			if err != nil {
				f.Fatalf("read %s: %v", e.Name(), err)
			}
			f.Add(data)
		}
	}
	f.Fuzz(func(t *testing.T, data []byte) {
		file, err := OpenBytes(data)
		if err != nil {
			return
		}
		file.WriteToBuffer()
	})
}

func FuzzParseContentXML(f *testing.F) {
	seeds := loadContentXMLSeeds(f)
	for _, seed := range seeds {
		f.Add(seed)
	}
	f.Fuzz(func(t *testing.T, data []byte) {
		file := NewFile()
		parseContentXML(file, data)
	})
}

func FuzzEvaluate(f *testing.F) {
	seeds := []string{
		"SUM(1;2;3)",
		"IF(1>0;\"yes\";\"no\")",
		"AVERAGE([.A1:.A5])",
		"VLOOKUP(\"x\";[.A1:.B3];2;0)",
		"[.A1]+[.B1]*[.C1]",
		"CONCATENATE(\"a\";\"b\";\"c\")",
		"ABS(-42)",
		"ROUND(3.14159;2)",
		"LEN(\"hello\")",
		"UPPER(\"hello\")",
		"MOD(10;3)",
		"POWER(2;10)",
		"IF(AND(1>0;2>1);\"both\";\"neither\")",
		"IFERROR(1/0;\"error\")",
	}
	for _, s := range seeds {
		f.Add(s)
	}
	f.Fuzz(func(t *testing.T, formula string) {
		values := CellValues{
			"A1": 1.0, "A2": 2.0, "A3": 3.0, "A4": 4.0, "A5": 5.0,
			"B1": 10.0, "B2": 20.0, "B3": 30.0,
			"C1": 100.0,
		}
		Evaluate(formula, values)
	})
}

func loadContentXMLSeeds(f *testing.F) [][]byte {
	f.Helper()
	entries, err := os.ReadDir("testdata")
	if err != nil {
		f.Fatalf("read testdata: %v", err)
	}
	var seeds [][]byte
	for _, e := range entries {
		if !strings.HasSuffix(e.Name(), ".ods") {
			continue
		}
		data, err := os.ReadFile(filepath.Join("testdata", e.Name()))
		if err != nil {
			continue
		}
		content, err := extractContentXMLFromODS(data)
		if err != nil {
			continue
		}
		seeds = append(seeds, content)
	}
	return seeds
}

func extractContentXMLFromODS(data []byte) ([]byte, error) {
	r, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return nil, err
	}
	for _, f := range r.File {
		if f.Name == "content.xml" {
			rc, err := f.Open()
			if err != nil {
				return nil, err
			}
			defer rc.Close()
			return io.ReadAll(rc)
		}
	}
	return nil, fmt.Errorf("content.xml not found")
}
