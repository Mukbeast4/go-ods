package xml

import "encoding/xml"

type DocumentStyles struct {
	XMLName         xml.Name        `xml:"document-styles"`
	Styles          OfficeStyles    `xml:"styles"`
	AutomaticStyles AutomaticStyles `xml:"automatic-styles"`
	MasterStyles    MasterStyles    `xml:"master-styles"`
}

type OfficeStyles struct {
	Styles        []Style        `xml:"style"`
	DefaultStyles []DefaultStyle `xml:"default-style"`
}

type DefaultStyle struct {
	Family              string               `xml:"family,attr,omitempty"`
	ParagraphProperties *ParagraphProperties `xml:"paragraph-properties,omitempty"`
	TextProperties      *TextProperties      `xml:"text-properties,omitempty"`
	TableCellProperties *TableCellProperties `xml:"table-cell-properties,omitempty"`
}

type MasterStyles struct {
	MasterPages []MasterPage `xml:"master-page"`
}

type MasterPage struct {
	Name           string `xml:"name,attr,omitempty"`
	PageLayoutName string `xml:"page-layout-name,attr,omitempty"`
}

func DefaultDocumentStyles() DocumentStyles {
	return DocumentStyles{
		Styles: OfficeStyles{
			DefaultStyles: []DefaultStyle{
				{
					Family: "table-cell",
					ParagraphProperties: &ParagraphProperties{
						TextAlign: "start",
					},
					TextProperties: &TextProperties{
						FontName: "Arial",
						FontSize: "10pt",
					},
				},
			},
		},
		MasterStyles: MasterStyles{
			MasterPages: []MasterPage{
				{Name: "Default", PageLayoutName: "pm1"},
			},
		},
	}
}
