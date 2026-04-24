package goods

import (
	"path/filepath"
	"strings"
	"testing"
)

func TestStyleRotation(t *testing.T) {
	f := NewFile()
	id, err := f.NewStyle(&Style{
		Alignment: &Alignment{Rotation: 90},
	})
	if err != nil {
		t.Fatalf("NewStyle: %v", err)
	}
	f.SetCellValue("Sheet1", "A1", "rotated")
	f.SetCellStyle("Sheet1", "A1", "A1", id)

	xml := contentOf(t, f)
	if !strings.Contains(xml, `style:rotation-angle="90"`) {
		t.Errorf("want rotation-angle=90 in XML, got:\n%s", xml)
	}
}

func TestStyleRotationNormalization(t *testing.T) {
	f := NewFile()
	id, _ := f.NewStyle(&Style{Alignment: &Alignment{Rotation: -90}})
	f.SetCellStyle("Sheet1", "A1", "A1", id)
	xml := contentOf(t, f)
	if !strings.Contains(xml, `style:rotation-angle="270"`) {
		t.Errorf("want rotation normalized to 270, got:\n%s", xml)
	}
}

func TestStyleIndent(t *testing.T) {
	f := NewFile()
	id, _ := f.NewStyle(&Style{
		Alignment: &Alignment{Indent: 3},
	})
	f.SetCellStyle("Sheet1", "A1", "A1", id)

	xml := contentOf(t, f)
	if !strings.Contains(xml, `fo:margin-left="0.7500cm"`) {
		t.Errorf("want margin-left=0.7500cm for indent=3, got:\n%s", xml)
	}
}

func TestStyleVerticalAlignSuperSub(t *testing.T) {
	f := NewFile()
	superID, _ := f.NewStyle(&Style{Font: &Font{VerticalAlign: "super"}})
	subID, _ := f.NewStyle(&Style{Font: &Font{VerticalAlign: "sub"}})

	f.SetCellStyle("Sheet1", "A1", "A1", superID)
	f.SetCellStyle("Sheet1", "B1", "B1", subID)

	xml := contentOf(t, f)
	if !strings.Contains(xml, `style:text-position="super 58%"`) {
		t.Errorf("want text-position super in XML, got:\n%s", xml)
	}
	if !strings.Contains(xml, `style:text-position="sub 58%"`) {
		t.Errorf("want text-position sub in XML, got:\n%s", xml)
	}
}

func TestStyleStrikethroughColor(t *testing.T) {
	f := NewFile()
	id, _ := f.NewStyle(&Style{
		Font: &Font{Strikethrough: true, StrikethroughColor: "#FF0000"},
	})
	f.SetCellStyle("Sheet1", "A1", "A1", id)

	xml := contentOf(t, f)
	if !strings.Contains(xml, `style:text-line-through-color="#FF0000"`) {
		t.Errorf("want text-line-through-color=#FF0000 in XML, got:\n%s", xml)
	}
}

func TestStrikethroughColorImpliesSolid(t *testing.T) {
	f := NewFile()
	id, _ := f.NewStyle(&Style{
		Font: &Font{StrikethroughColor: "#00FF00"},
	})
	f.SetCellStyle("Sheet1", "A1", "A1", id)

	xml := contentOf(t, f)
	if !strings.Contains(xml, `style:text-line-through-style="solid"`) {
		t.Errorf("want text-line-through-style=solid when only color is set, got:\n%s", xml)
	}
}

func TestRowAutoFit(t *testing.T) {
	f := NewFile()
	if err := f.SetRowAutoFit("Sheet1", 3, true); err != nil {
		t.Fatalf("SetRowAutoFit: %v", err)
	}

	got, err := f.GetRowAutoFit("Sheet1", 3)
	if err != nil {
		t.Fatalf("GetRowAutoFit: %v", err)
	}
	if !got {
		t.Error("want autoFit=true")
	}

	xml := contentOf(t, f)
	if !strings.Contains(xml, `style:use-optimal-row-height="true"`) {
		t.Errorf("want use-optimal-row-height=true in XML, got:\n%s", xml)
	}
}

func TestRowAutoFitRoundTrip(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "rf.ods")

	src := NewFile()
	src.SetRowAutoFit("Sheet1", 5, true)
	if err := src.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	dst, err := OpenFile(path)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}
	defer dst.Close()

	got, _ := dst.GetRowAutoFit("Sheet1", 5)
	if !got {
		t.Error("row autoFit lost on roundtrip")
	}
}

func TestStyleAdvancedRoundTrip(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "style.ods")

	src := NewFile()
	id, _ := src.NewStyle(&Style{
		Font: &Font{
			Family:             "Arial",
			Size:               "12pt",
			Strikethrough:      true,
			StrikethroughColor: "#FF0000",
			VerticalAlign:      "super",
		},
		Alignment: &Alignment{
			Horizontal: "center",
			Rotation:   45,
			Indent:     2,
		},
	})
	src.SetCellValue("Sheet1", "A1", "styled")
	src.SetCellStyle("Sheet1", "A1", "A1", id)

	if err := src.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	dst, err := OpenFile(path)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}
	defer dst.Close()

	val, _ := dst.GetCellValue("Sheet1", "A1")
	if val != "styled" {
		t.Errorf("value = %q, want styled", val)
	}
}

func contentOf(t *testing.T, f *File) string {
	t.Helper()
	buf, err := f.WriteToBuffer()
	if err != nil {
		t.Fatalf("WriteToBuffer: %v", err)
	}
	return string(extractContentXMLBytes(t, buf.Bytes()))
}
