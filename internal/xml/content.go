package xml

import "encoding/xml"

type DocumentContent struct {
	XMLName         xml.Name        `xml:"document-content"`
	AutomaticStyles AutomaticStyles `xml:"automatic-styles"`
	Body            Body            `xml:"body"`
}

type AutomaticStyles struct {
	Styles []Style `xml:"style"`
}

type Body struct {
	Spreadsheet Spreadsheet `xml:"spreadsheet"`
}

type Spreadsheet struct {
	Tables []Table `xml:"table"`
}

type Table struct {
	Name        string        `xml:"name,attr"`
	PrintRanges string        `xml:"print-ranges,attr,omitempty"`
	Columns     []TableColumn `xml:"table-column"`
	Rows        []TableRow    `xml:"table-row"`
}

type TableColumn struct {
	StyleName             string `xml:"style-name,attr,omitempty"`
	DefaultCellStyleName  string `xml:"default-cell-style-name,attr,omitempty"`
	NumberColumnsRepeated int    `xml:"number-columns-repeated,attr,omitempty"`
}

type TableRow struct {
	StyleName          string      `xml:"style-name,attr,omitempty"`
	NumberRowsRepeated int         `xml:"number-rows-repeated,attr,omitempty"`
	Cells              []TableCell `xml:"table-cell"`
}

type TableCell struct {
	ValueType              string      `xml:"value-type,attr,omitempty"`
	Value                  string      `xml:"value,attr,omitempty"`
	DateValue              string      `xml:"date-value,attr,omitempty"`
	BooleanValue           string      `xml:"boolean-value,attr,omitempty"`
	StringValue            string      `xml:"string-value,attr,omitempty"`
	Formula                string      `xml:"formula,attr,omitempty"`
	StyleName              string      `xml:"style-name,attr,omitempty"`
	ContentValidationName  string      `xml:"content-validation-name,attr,omitempty"`
	NumberColumnsRepeated  int         `xml:"number-columns-repeated,attr,omitempty"`
	NumberColumnsSpanned   int         `xml:"number-columns-spanned,attr,omitempty"`
	NumberRowsSpanned      int         `xml:"number-rows-spanned,attr,omitempty"`
	Annotation             *Annotation `xml:"annotation,omitempty"`
	Paragraphs             []TextP     `xml:"p"`
}

type Annotation struct {
	Creator string
	Date    string
	Text    string
}

type NamedRange struct {
	Name             string
	BaseCellAddress  string
	CellRangeAddress string
}

type ContentValidation struct {
	Name         string
	Condition    string
	AllowEmpty   bool
	ErrorMessage *ErrorMessage
	HelpMessage  *HelpMessage
}

type ErrorMessage struct {
	Display     bool
	MessageType string
	Title       string
	Text        string
}

type HelpMessage struct {
	Display bool
	Title   string
	Text    string
}

type DatabaseRange struct {
	Name                 string
	TargetRangeAddress   string
	DisplayFilterButtons string
}

type TextP struct {
	Text string `xml:",chardata"`
	Link *TextA `xml:"a,omitempty"`
}

type TextA struct {
	Href string `xml:"href,attr,omitempty"`
	Type string `xml:"type,attr,omitempty"`
	Text string `xml:",chardata"`
}

type Style struct {
	Name                  string                 `xml:"name,attr,omitempty"`
	Family                string                 `xml:"family,attr,omitempty"`
	ParentStyleName       string                 `xml:"parent-style-name,attr,omitempty"`
	DataStyleName         string                 `xml:"data-style-name,attr,omitempty"`
	TableProperties       *TableProperties       `xml:"table-properties,omitempty"`
	TableRowProperties    *TableRowProperties    `xml:"table-row-properties,omitempty"`
	TableColumnProperties *TableColumnProperties `xml:"table-column-properties,omitempty"`
	TableCellProperties   *TableCellProperties   `xml:"table-cell-properties,omitempty"`
	TextProperties        *TextProperties        `xml:"text-properties,omitempty"`
	ParagraphProperties   *ParagraphProperties   `xml:"paragraph-properties,omitempty"`
}

type TableProperties struct {
	Display string `xml:"display,attr,omitempty"`
}

type TableRowProperties struct {
	RowHeight        string `xml:"row-height,attr,omitempty"`
	UseOptimalHeight string `xml:"use-optimal-row-height,attr,omitempty"`
}

type TableColumnProperties struct {
	ColumnWidth     string `xml:"column-width,attr,omitempty"`
	UseOptimalWidth string `xml:"use-optimal-column-width,attr,omitempty"`
}

type TableCellProperties struct {
	BackgroundColor string `xml:"background-color,attr,omitempty"`
	BorderTop       string `xml:"border-top,attr,omitempty"`
	BorderBottom    string `xml:"border-bottom,attr,omitempty"`
	BorderLeft      string `xml:"border-left,attr,omitempty"`
	BorderRight     string `xml:"border-right,attr,omitempty"`
	Border          string `xml:"border,attr,omitempty"`
	VerticalAlign   string `xml:"vertical-align,attr,omitempty"`
	WrapOption      string `xml:"wrap-option,attr,omitempty"`
}

type TextProperties struct {
	FontName             string `xml:"font-name,attr,omitempty"`
	FontSize             string `xml:"font-size,attr,omitempty"`
	FontWeight           string `xml:"font-weight,attr,omitempty"`
	FontStyle            string `xml:"font-style,attr,omitempty"`
	Color                string `xml:"color,attr,omitempty"`
	TextUnderlineStyle   string `xml:"text-underline-style,attr,omitempty"`
	TextLineThroughStyle string `xml:"text-line-through-style,attr,omitempty"`
}

type ParagraphProperties struct {
	TextAlign string `xml:"text-align,attr,omitempty"`
}
