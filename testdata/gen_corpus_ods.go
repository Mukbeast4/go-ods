//go:build ignore

package main

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"log"
	"os"
)

const mimeTypeODS = "application/vnd.oasis.opendocument.spreadsheet"

var contentXMLHeader = xml.Header + `<office:document-content` +
	` xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"` +
	` xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"` +
	` xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"` +
	` xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"` +
	` xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"` +
	` xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"` +
	` xmlns:xlink="http://www.w3.org/1999/xlink"` +
	` xmlns:dc="http://purl.org/dc/elements/1.1/"` +
	` xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"` +
	` xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"` +
	` xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"` +
	` xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"` +
	` xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0"` +
	` xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0"` +
	` xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"` +
	` office:version="1.2">`

const contentXMLFooter = `</office:document-content>`

const minimalStylesXML = xml.Header + `<office:document-styles` +
	` xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"` +
	` xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"` +
	` office:version="1.2">` +
	`<office:styles/>` +
	`<office:automatic-styles/>` +
	`<office:master-styles><style:master-page style:name="Default" style:page-layout-name="pm1"/></office:master-styles>` +
	`</office:document-styles>`

const minimalManifestXML = xml.Header + `<manifest:manifest` +
	` xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"` +
	` manifest:version="1.2">` +
	`<manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet" manifest:version="1.2"/>` +
	`<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>` +
	`<manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>` +
	`</manifest:manifest>`

func main() {
	generators := []struct {
		name string
		fn   func() error
	}{
		{"corpus_repeated_cells.ods", genCorpusRepeatedCells},
		{"corpus_empty_rows.ods", genCorpusEmptyRows},
		{"corpus_nested_styles.ods", genCorpusNestedStyles},
		{"corpus_special_chars.ods", genCorpusSpecialChars},
	}

	for _, g := range generators {
		if err := g.fn(); err != nil {
			log.Fatalf("generate %s: %v", g.name, err)
		}
		log.Printf("generated testdata/%s", g.name)
	}
}

func buildODSFromContentXML(contentXML string) ([]byte, error) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)

	mt, err := zw.Create("mimetype")
	if err != nil {
		return nil, err
	}
	mt.Write([]byte(mimeTypeODS))

	cw, err := zw.Create("content.xml")
	if err != nil {
		return nil, err
	}
	cw.Write([]byte(contentXML))

	sw, err := zw.Create("styles.xml")
	if err != nil {
		return nil, err
	}
	sw.Write([]byte(minimalStylesXML))

	mw, err := zw.Create("META-INF/manifest.xml")
	if err != nil {
		return nil, err
	}
	mw.Write([]byte(minimalManifestXML))

	if err := zw.Close(); err != nil {
		return nil, err
	}

	return buf.Bytes(), nil
}

func writeODS(path, contentBody string) error {
	contentXML := contentXMLHeader +
		`<office:automatic-styles/>` +
		`<office:body><office:spreadsheet>` +
		contentBody +
		`</office:spreadsheet></office:body>` +
		contentXMLFooter

	data, err := buildODSFromContentXML(contentXML)
	if err != nil {
		return err
	}
	return os.WriteFile(path, data, 0644)
}

func writeODSWithStyles(path, stylesBody, contentBody string) error {
	contentXML := contentXMLHeader +
		`<office:automatic-styles>` + stylesBody + `</office:automatic-styles>` +
		`<office:body><office:spreadsheet>` +
		contentBody +
		`</office:spreadsheet></office:body>` +
		contentXMLFooter

	data, err := buildODSFromContentXML(contentXML)
	if err != nil {
		return err
	}
	return os.WriteFile(path, data, 0644)
}

func genCorpusRepeatedCells() error {
	body := `<table:table table:name="Sheet1">` +
		`<table:table-column/>` +
		`<table:table-row>` +
		`<table:table-cell office:value-type="float" office:value="42" table:number-columns-repeated="5">` +
		`<text:p>42</text:p>` +
		`</table:table-cell>` +
		`<table:table-cell office:value-type="string" office:string-value="end">` +
		`<text:p>end</text:p>` +
		`</table:table-cell>` +
		`</table:table-row>` +
		`<table:table-row table:number-rows-repeated="3">` +
		`<table:table-cell office:value-type="string" office:string-value="repeated_row">` +
		`<text:p>repeated_row</text:p>` +
		`</table:table-cell>` +
		`</table:table-row>` +
		`<table:table-row>` +
		`<table:table-cell office:value-type="string" office:string-value="after_repeated">` +
		`<text:p>after_repeated</text:p>` +
		`</table:table-cell>` +
		`</table:table-row>` +
		`</table:table>`

	return writeODS("testdata/corpus_repeated_cells.ods", body)
}

func genCorpusEmptyRows() error {
	body := `<table:table table:name="Sheet1">` +
		`<table:table-column/>` +
		`<table:table-row>` +
		`<table:table-cell office:value-type="string" office:string-value="header">` +
		`<text:p>header</text:p>` +
		`</table:table-cell>` +
		`</table:table-row>` +
		`<table:table-row>` +
		`<table:table-cell office:value-type="float" office:value="100">` +
		`<text:p>100</text:p>` +
		`</table:table-cell>` +
		`</table:table-row>` +
		`<table:table-row table:number-rows-repeated="65534">` +
		`<table:table-cell/>` +
		`</table:table-row>` +
		`</table:table>`

	return writeODS("testdata/corpus_empty_rows.ods", body)
}

func genCorpusNestedStyles() error {
	styles := `<style:style style:name="ce-parent" style:family="table-cell">` +
		`<style:text-properties fo:font-weight="bold"/>` +
		`</style:style>` +
		`<style:style style:name="ce-child" style:family="table-cell" style:parent-style-name="ce-parent" style:data-style-name="N100">` +
		`<style:text-properties fo:color="#FF0000"/>` +
		`</style:style>` +
		`<style:style style:name="ce-grandchild" style:family="table-cell" style:parent-style-name="ce-child">` +
		`<style:text-properties fo:font-style="italic"/>` +
		`</style:style>`

	body := `<table:table table:name="Sheet1">` +
		`<table:table-column/>` +
		`<table:table-row>` +
		`<table:table-cell table:style-name="ce-parent" office:value-type="string" office:string-value="parent">` +
		`<text:p>parent</text:p>` +
		`</table:table-cell>` +
		`</table:table-row>` +
		`<table:table-row>` +
		`<table:table-cell table:style-name="ce-child" office:value-type="string" office:string-value="child">` +
		`<text:p>child</text:p>` +
		`</table:table-cell>` +
		`</table:table-row>` +
		`<table:table-row>` +
		`<table:table-cell table:style-name="ce-grandchild" office:value-type="string" office:string-value="grandchild">` +
		`<text:p>grandchild</text:p>` +
		`</table:table-cell>` +
		`</table:table-row>` +
		`</table:table>`

	return writeODSWithStyles("testdata/corpus_nested_styles.ods", styles, body)
}

func genCorpusSpecialChars() error {
	escape := func(s string) string {
		var buf bytes.Buffer
		xml.EscapeText(&buf, []byte(s))
		return buf.String()
	}

	type testCase struct {
		label string
		value string
	}

	cases := []testCase{
		{"lt", "<less than>"},
		{"gt", "greater>than"},
		{"amp", "this & that"},
		{"quot", `say "hello"`},
		{"apos", "it's here"},
		{"newline", "line1\nline2\nline3"},
		{"tab", "col1\tcol2"},
		{"cjk", "\u4f60\u597d\u4e16\u754c"},
		{"emoji", "\U0001F600\U0001F680\U0001F4A1"},
		{"rtl", "\u0645\u0631\u062D\u0628\u0627"},
		{"mixed", "Hello <World> & \"Friends\" \u2603"},
		{"empty", ""},
		{"spaces", "  leading and trailing  "},
		{"unicode_math", "\u222B\u2211\u221E\u00B1"},
	}

	var rows string
	for _, tc := range cases {
		rows += `<table:table-row>` +
			`<table:table-cell office:value-type="string" office:string-value="` + escape(tc.label) + `">` +
			`<text:p>` + escape(tc.label) + `</text:p>` +
			`</table:table-cell>` +
			`<table:table-cell office:value-type="string" office:string-value="` + escape(tc.value) + `">` +
			`<text:p>` + escape(tc.value) + `</text:p>` +
			`</table:table-cell>` +
			`</table:table-row>`
	}

	body := `<table:table table:name="Sheet1">` +
		`<table:table-column/><table:table-column/>` +
		rows +
		`</table:table>`

	return writeODS("testdata/corpus_special_chars.ods", body)
}
