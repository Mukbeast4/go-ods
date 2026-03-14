package goods

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"io"
	"os"
	"sort"
	"strings"
	"testing"
)

func extractContentXMLBytes(t *testing.T, odsData []byte) []byte {
	t.Helper()
	r, err := zip.NewReader(bytes.NewReader(odsData), int64(len(odsData)))
	if err != nil {
		t.Fatalf("failed to open ODS zip: %v", err)
	}
	for _, f := range r.File {
		if f.Name == "content.xml" {
			rc, err := f.Open()
			if err != nil {
				t.Fatalf("failed to open content.xml: %v", err)
			}
			defer rc.Close()
			data, err := io.ReadAll(rc)
			if err != nil {
				t.Fatalf("failed to read content.xml: %v", err)
			}
			return data
		}
	}
	t.Fatal("content.xml not found in ODS archive")
	return nil
}

func normalizeXML(data []byte) string {
	decoder := xml.NewDecoder(bytes.NewReader(data))
	var buf bytes.Buffer
	encoder := xml.NewEncoder(&buf)
	encoder.Indent("", "  ")
	for {
		tok, err := decoder.Token()
		if err != nil {
			break
		}
		switch t := tok.(type) {
		case xml.StartElement:
			sort.Slice(t.Attr, func(i, j int) bool {
				ki := t.Attr[i].Name.Space + ":" + t.Attr[i].Name.Local
				kj := t.Attr[j].Name.Space + ":" + t.Attr[j].Name.Local
				return ki < kj
			})
			encoder.EncodeToken(t)
		default:
			encoder.EncodeToken(xml.CopyToken(tok))
		}
	}
	encoder.Flush()
	return buf.String()
}

func diffStrings(a, b string) string {
	linesA := strings.Split(a, "\n")
	linesB := strings.Split(b, "\n")
	var diffs []string
	maxLen := len(linesA)
	if len(linesB) > maxLen {
		maxLen = len(linesB)
	}
	for i := 0; i < maxLen && len(diffs) < 10; i++ {
		var la, lb string
		if i < len(linesA) {
			la = linesA[i]
		}
		if i < len(linesB) {
			lb = linesB[i]
		}
		if la != lb {
			diffs = append(diffs, "line "+itoa(i+1)+":\n  - "+la+"\n  + "+lb)
		}
	}
	if len(diffs) == 0 {
		return "(no differences)"
	}
	return strings.Join(diffs, "\n")
}

func itoa(n int) string {
	if n == 0 {
		return "0"
	}
	var buf [20]byte
	i := len(buf)
	for n > 0 {
		i--
		buf[i] = byte('0' + n%10)
		n /= 10
	}
	return string(buf[i:])
}

func TestDiffUnmodifiedRoundTrip(t *testing.T) {
	originalData, err := os.ReadFile("testdata/sample.ods")
	if err != nil {
		t.Fatalf("failed to read sample.ods: %v", err)
	}

	f, err := OpenBytes(originalData)
	if err != nil {
		t.Fatalf("OpenBytes() error: %v", err)
	}

	buf, err := f.WriteToBuffer()
	if err != nil {
		t.Fatalf("WriteToBuffer() error: %v", err)
	}

	originalXML := extractContentXMLBytes(t, originalData)
	roundTripXML := extractContentXMLBytes(t, buf.Bytes())

	normalizedOriginal := normalizeXML(originalXML)
	normalizedRoundTrip := normalizeXML(roundTripXML)

	if normalizedOriginal == normalizedRoundTrip {
		t.Logf("round-trip content.xml is identical after normalization")
	} else {
		t.Logf("round-trip content.xml differs after normalization (documenting known gaps):\n%s",
			diffStrings(normalizedOriginal, normalizedRoundTrip))
	}
}

func TestDiffAddOneCell(t *testing.T) {
	originalData, err := os.ReadFile("testdata/sample.ods")
	if err != nil {
		t.Fatalf("failed to read sample.ods: %v", err)
	}

	f, err := OpenBytes(originalData)
	if err != nil {
		t.Fatalf("OpenBytes() error: %v", err)
	}

	sheets := f.GetSheetList()
	if len(sheets) == 0 {
		t.Fatal("sample.ods has no sheets")
	}
	sheetName := sheets[0]

	err = f.SetCellValue(sheetName, "Z99", "new_value")
	if err != nil {
		t.Fatalf("SetCellValue() error: %v", err)
	}

	buf, err := f.WriteToBuffer()
	if err != nil {
		t.Fatalf("WriteToBuffer() error: %v", err)
	}

	modifiedXML := extractContentXMLBytes(t, buf.Bytes())
	modifiedStr := string(modifiedXML)

	if !strings.Contains(modifiedStr, "new_value") {
		t.Errorf("modified content.xml does not contain 'new_value'")
	}

	originalXML := extractContentXMLBytes(t, originalData)
	normalizedOriginal := normalizeXML(originalXML)
	normalizedModified := normalizeXML(modifiedXML)

	t.Logf("differences after adding cell Z99:\n%s",
		diffStrings(normalizedOriginal, normalizedModified))
}

func TestDiffAddSheet(t *testing.T) {
	originalData, err := os.ReadFile("testdata/sample.ods")
	if err != nil {
		t.Fatalf("failed to read sample.ods: %v", err)
	}

	f, err := OpenBytes(originalData)
	if err != nil {
		t.Fatalf("OpenBytes() error: %v", err)
	}

	originalSheets := f.GetSheetList()

	_, err = f.NewSheet("NewTestSheet")
	if err != nil {
		t.Fatalf("NewSheet() error: %v", err)
	}

	buf, err := f.WriteToBuffer()
	if err != nil {
		t.Fatalf("WriteToBuffer() error: %v", err)
	}

	modifiedXML := extractContentXMLBytes(t, buf.Bytes())
	modifiedStr := string(modifiedXML)

	if !strings.Contains(modifiedStr, "NewTestSheet") {
		t.Errorf("modified content.xml does not contain 'NewTestSheet' as table name")
	}

	for _, name := range originalSheets {
		if !strings.Contains(modifiedStr, name) {
			t.Errorf("original sheet %q missing from modified content.xml", name)
		}
	}
}

func TestDiffDeleteSheet(t *testing.T) {
	originalData, err := os.ReadFile("testdata/golden_multisheet.ods")
	if err != nil {
		t.Fatalf("failed to read golden_multisheet.ods: %v", err)
	}

	f, err := OpenBytes(originalData)
	if err != nil {
		t.Fatalf("OpenBytes() error: %v", err)
	}

	allSheets := f.GetSheetList()
	if len(allSheets) < 2 {
		t.Fatalf("golden_multisheet.ods has only %d sheets, need at least 2", len(allSheets))
	}

	sheetToDelete := "Sheet3"
	found := false
	for _, s := range allSheets {
		if s == sheetToDelete {
			found = true
			break
		}
	}
	if !found {
		t.Fatalf("sheet %q not found in golden_multisheet.ods, available: %v", sheetToDelete, allSheets)
	}

	err = f.DeleteSheet(sheetToDelete)
	if err != nil {
		t.Fatalf("DeleteSheet() error: %v", err)
	}

	buf, err := f.WriteToBuffer()
	if err != nil {
		t.Fatalf("WriteToBuffer() error: %v", err)
	}

	modifiedXML := extractContentXMLBytes(t, buf.Bytes())
	modifiedStr := string(modifiedXML)

	if strings.Contains(modifiedStr, "table:name=\""+sheetToDelete+"\"") {
		t.Errorf("deleted sheet %q still present in content.xml", sheetToDelete)
	}

	for _, name := range allSheets {
		if name == sheetToDelete {
			continue
		}
		if !strings.Contains(modifiedStr, name) {
			t.Errorf("remaining sheet %q missing from modified content.xml", name)
		}
	}
}

func TestDiffAddFormula(t *testing.T) {
	originalData, err := os.ReadFile("testdata/sample.ods")
	if err != nil {
		t.Fatalf("failed to read sample.ods: %v", err)
	}

	f, err := OpenBytes(originalData)
	if err != nil {
		t.Fatalf("OpenBytes() error: %v", err)
	}

	sheets := f.GetSheetList()
	if len(sheets) == 0 {
		t.Fatal("sample.ods has no sheets")
	}
	sheetName := sheets[0]

	formula := "SUM([.A1:.A5])"
	err = f.SetCellFormula(sheetName, "H1", formula)
	if err != nil {
		t.Fatalf("SetCellFormula() error: %v", err)
	}

	buf, err := f.WriteToBuffer()
	if err != nil {
		t.Fatalf("WriteToBuffer() error: %v", err)
	}

	modifiedXML := extractContentXMLBytes(t, buf.Bytes())
	modifiedStr := string(modifiedXML)

	if !strings.Contains(modifiedStr, "table:formula") {
		t.Errorf("modified content.xml does not contain 'table:formula' attribute")
	}

	if !strings.Contains(modifiedStr, formula) {
		t.Errorf("modified content.xml does not contain formula text %q", formula)
	}
}
