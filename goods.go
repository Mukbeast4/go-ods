package goods

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"strings"
	"time"

	oxml "github.com/mukbeast4/go-ods/internal/xml"
	ozip "github.com/mukbeast4/go-ods/internal/zip"
)

type File struct {
	sheets      []*sheet
	activeSheet int
	styles      *styleManager
	metadata    *oxml.DocumentMeta
	docStyles   *oxml.DocumentStyles
	rawFiles    map[string][]byte
	path        string
	closed      bool
	autoRecalc  bool
}

type sheet struct {
	name    string
	columns []column
	rows    map[int]*row
	merges  []mergeRange
	maxRow  int
	maxCol  int
}

type column struct {
	width   float64
	styleID int
}

type row struct {
	cells   map[int]*cell
	height  float64
	visible bool
}

type cell struct {
	valueType CellType
	rawValue  string
	formula   string
	styleID   int
	colSpan   int
	rowSpan   int
}

type mergeRange struct {
	startCol, startRow int
	endCol, endRow     int
}

func NewFile() *File {
	f := &File{
		sheets:   make([]*sheet, 0),
		styles:   newStyleManager(),
		rawFiles: make(map[string][]byte),
		metadata: &oxml.DocumentMeta{
			Meta: oxml.Meta{
				Generator:    "goods",
				CreationDate: time.Now().UTC().Format("2006-01-02T15:04:05"),
				Date:         time.Now().UTC().Format("2006-01-02T15:04:05"),
			},
		},
		docStyles: defaultDocStyles(),
	}

	s := &sheet{
		name:    "Sheet1",
		rows:    make(map[int]*row),
		columns: make([]column, 0),
	}
	f.sheets = append(f.sheets, s)

	return f
}

func OpenFile(path string) (*File, error) {
	result, err := ozip.ReadFile(path)
	if err != nil {
		return nil, fmt.Errorf("open file: %w", err)
	}

	f, err := parseZipResult(result)
	if err != nil {
		return nil, err
	}
	f.path = path
	return f, nil
}

func OpenReader(r io.ReaderAt, size int64) (*File, error) {
	result, err := ozip.ReadFromReader(r, size)
	if err != nil {
		return nil, fmt.Errorf("open reader: %w", err)
	}
	return parseZipResult(result)
}

func OpenBytes(data []byte) (*File, error) {
	result, err := ozip.ReadBytes(data)
	if err != nil {
		return nil, fmt.Errorf("open bytes: %w", err)
	}
	return parseZipResult(result)
}

func parseZipResult(result *ozip.ReadResult) (*File, error) {
	f := &File{
		sheets:   make([]*sheet, 0),
		styles:   newStyleManager(),
		rawFiles: make(map[string][]byte),
		metadata: &oxml.DocumentMeta{
			Meta: oxml.Meta{Generator: "goods"},
		},
		docStyles: defaultDocStyles(),
	}

	for name, data := range result.Files {
		switch name {
		case "mimetype", "content.xml", "styles.xml", "meta.xml", "META-INF/manifest.xml":
		default:
			f.rawFiles[name] = data
		}
	}

	contentData := result.Files["content.xml"]
	if err := parseContentXML(f, contentData); err != nil {
		return nil, fmt.Errorf("parse content.xml: %w", err)
	}

	if metaData, ok := result.Files["meta.xml"]; ok {
		parseMetaXML(f, metaData)
	}

	if stylesData, ok := result.Files["styles.xml"]; ok {
		parseStylesXML(f, stylesData)
	}

	return f, nil
}

type xmlParagraph struct {
	Text string `xml:",chardata"`
}

type xmlTableCell struct {
	XMLName               xml.Name       `xml:"table-cell"`
	ValueType             string         `xml:"value-type,attr"`
	Value                 string         `xml:"value,attr"`
	DateValue             string         `xml:"date-value,attr"`
	BooleanValue          string         `xml:"boolean-value,attr"`
	StringValue           string         `xml:"string-value,attr"`
	Formula               string         `xml:"formula,attr"`
	StyleName             string         `xml:"style-name,attr"`
	NumberColumnsRepeated int            `xml:"number-columns-repeated,attr"`
	NumberColumnsSpanned  int            `xml:"number-columns-spanned,attr"`
	NumberRowsSpanned     int            `xml:"number-rows-spanned,attr"`
	Paragraphs            []xmlParagraph `xml:"p"`
}

type xmlTableRow struct {
	XMLName            xml.Name       `xml:"table-row"`
	StyleName          string         `xml:"style-name,attr"`
	NumberRowsRepeated int            `xml:"number-rows-repeated,attr"`
	Cells              []xmlTableCell `xml:"table-cell"`
}

type xmlTableColumn struct {
	XMLName               xml.Name `xml:"table-column"`
	StyleName             string   `xml:"style-name,attr"`
	NumberColumnsRepeated int      `xml:"number-columns-repeated,attr"`
}

type xmlTable struct {
	XMLName xml.Name         `xml:"table"`
	Name    string           `xml:"name,attr"`
	Columns []xmlTableColumn `xml:"table-column"`
	Rows    []xmlTableRow    `xml:"table-row"`
}

type xmlContent struct {
	XMLName xml.Name `xml:"document-content"`
	Body    struct {
		Spreadsheet struct {
			Tables []xmlTable `xml:"table"`
		} `xml:"spreadsheet"`
	} `xml:"body"`
}

func parseContentXML(f *File, data []byte) error {
	var content xmlContent
	if err := xml.Unmarshal(data, &content); err != nil {
		return fmt.Errorf("unmarshal content: %w", err)
	}

	for _, xmlTbl := range content.Body.Spreadsheet.Tables {
		s := &sheet{
			name:    xmlTbl.Name,
			rows:    make(map[int]*row),
			columns: make([]column, 0),
		}

		for _, xmlCol := range xmlTbl.Columns {
			count := xmlCol.NumberColumnsRepeated
			if count < 1 {
				count = 1
			}
			for range count {
				s.columns = append(s.columns, column{})
			}
		}

		rowIdx := 1
		for _, xmlRow := range xmlTbl.Rows {
			rowRepeat := xmlRow.NumberRowsRepeated
			if rowRepeat < 1 {
				rowRepeat = 1
			}

			if rowRepeat > 1000 && isEmptyRow(xmlRow.Cells) {
				rowIdx += rowRepeat
				continue
			}

			for range rowRepeat {
				colIdx := 1
				hasData := false

				for _, xmlCell := range xmlRow.Cells {
					colRepeat := xmlCell.NumberColumnsRepeated
					if colRepeat < 1 {
						colRepeat = 1
					}

					c := convertXMLCell(&xmlCell)
					if c != nil {
						hasData = true
						for rep := range colRepeat {
							r := s.getOrCreateRow(rowIdx)
							r.cells[colIdx+rep] = &cell{
								valueType: c.valueType,
								rawValue:  c.rawValue,
								formula:   c.formula,
								styleID:   c.styleID,
								colSpan:   c.colSpan,
								rowSpan:   c.rowSpan,
							}
							if colIdx+rep > s.maxCol {
								s.maxCol = colIdx + rep
							}
						}
					}
					colIdx += colRepeat
				}

				if hasData {
					if rowIdx > s.maxRow {
						s.maxRow = rowIdx
					}
				}
				rowIdx++
			}
		}

		f.sheets = append(f.sheets, s)
	}

	if len(f.sheets) == 0 {
		s := &sheet{
			name:    "Sheet1",
			rows:    make(map[int]*row),
			columns: make([]column, 0),
		}
		f.sheets = append(f.sheets, s)
	}

	return nil
}

func isEmptyRow(cells []xmlTableCell) bool {
	for _, c := range cells {
		if c.ValueType != "" || len(c.Paragraphs) > 0 || c.Formula != "" {
			return false
		}
	}
	return true
}

func convertXMLCell(xc *xmlTableCell) *cell {
	if xc.ValueType == "" && len(xc.Paragraphs) == 0 && xc.Formula == "" {
		return nil
	}

	c := &cell{
		valueType: cellTypeFromODS(xc.ValueType),
		colSpan:   xc.NumberColumnsSpanned,
		rowSpan:   xc.NumberRowsSpanned,
	}

	switch xc.ValueType {
	case "float", "currency", "percentage":
		c.rawValue = xc.Value
	case "date":
		c.rawValue = xc.DateValue
	case "boolean":
		c.rawValue = xc.BooleanValue
	case "string", "":
		texts := make([]string, 0, len(xc.Paragraphs))
		for _, p := range xc.Paragraphs {
			texts = append(texts, p.Text)
		}
		if len(texts) > 0 {
			c.rawValue = strings.Join(texts, "\n")
			if xc.ValueType == "" {
				c.valueType = CellTypeString
			}
		}
	}

	if xc.Formula != "" {
		c.formula = strings.TrimPrefix(xc.Formula, "of:=")
		c.formula = strings.TrimPrefix(c.formula, "of:")
	}

	return c
}

func parseMetaXML(f *File, data []byte) {
	type xmlMeta struct {
		XMLName xml.Name `xml:"document-meta"`
		Meta    struct {
			Generator    string `xml:"generator"`
			Title        string `xml:"title"`
			Description  string `xml:"description"`
			Subject      string `xml:"subject"`
			Creator      string `xml:"creator"`
			CreationDate string `xml:"creation-date"`
			Date         string `xml:"date"`
		} `xml:"meta"`
	}

	var m xmlMeta
	if err := xml.Unmarshal(data, &m); err != nil {
		return
	}

	f.metadata = &oxml.DocumentMeta{
		Meta: oxml.Meta{
			Generator:    m.Meta.Generator,
			Title:        m.Meta.Title,
			Description:  m.Meta.Description,
			Subject:      m.Meta.Subject,
			Creator:      m.Meta.Creator,
			CreationDate: m.Meta.CreationDate,
			Date:         m.Meta.Date,
		},
	}
}

func parseStylesXML(f *File, data []byte) {
	_ = data
}

func (s *sheet) getOrCreateRow(rowIdx int) *row {
	r, ok := s.rows[rowIdx]
	if !ok {
		r = &row{
			cells:   make(map[int]*cell),
			visible: true,
		}
		s.rows[rowIdx] = r
	}
	return r
}

func (f *File) getSheet(name string) *sheet {
	for _, s := range f.sheets {
		if s.name == name {
			return s
		}
	}
	return nil
}

func defaultDocStyles() *oxml.DocumentStyles {
	ds := oxml.DefaultDocumentStyles()
	return &ds
}

func (f *File) buildContentXML() (*oxml.DocumentContent, []oxml.Style) {
	doc := &oxml.DocumentContent{}
	var autoStyles []oxml.Style

	for _, s := range f.sheets {
		table := oxml.Table{
			Name: s.name,
		}

		colCount := s.maxCol
		if len(s.columns) > colCount {
			colCount = len(s.columns)
		}
		if colCount < 1 {
			colCount = 1
		}

		for i := range colCount {
			col := oxml.TableColumn{}
			if i < len(s.columns) && s.columns[i].width > 0 {
				styleName := fmt.Sprintf("co%d", len(autoStyles)+1)
				autoStyles = append(autoStyles, oxml.Style{
					Name:   styleName,
					Family: "table-column",
					TableColumnProperties: &oxml.TableColumnProperties{
						ColumnWidth: fmt.Sprintf("%.4fcm", s.columns[i].width),
					},
				})
				col.StyleName = styleName
			}
			table.Columns = append(table.Columns, col)
		}

		maxRow := s.maxRow
		if maxRow < 1 {
			maxRow = 1
		}

		for rowIdx := 1; rowIdx <= maxRow; rowIdx++ {
			r, exists := s.rows[rowIdx]
			xmlRow := oxml.TableRow{}

			if exists && r.height > 0 {
				styleName := fmt.Sprintf("ro%d", len(autoStyles)+1)
				autoStyles = append(autoStyles, oxml.Style{
					Name:   styleName,
					Family: "table-row",
					TableRowProperties: &oxml.TableRowProperties{
						RowHeight:        fmt.Sprintf("%.4fcm", r.height),
						UseOptimalHeight: "false",
					},
				})
				xmlRow.StyleName = styleName
			}

			for colIdx := 1; colIdx <= colCount; colIdx++ {
				xmlCell := oxml.TableCell{}

				if exists {
					if c, ok := r.cells[colIdx]; ok {
						xmlCell = buildXMLCell(c, f.styles, &autoStyles)
					}
				}

				xmlRow.Cells = append(xmlRow.Cells, xmlCell)
			}

			table.Rows = append(table.Rows, xmlRow)
		}

		doc.Body.Spreadsheet.Tables = append(doc.Body.Spreadsheet.Tables, table)
	}

	return doc, autoStyles
}

func buildXMLCell(c *cell, sm *styleManager, autoStyles *[]oxml.Style) oxml.TableCell {
	xmlCell := oxml.TableCell{}

	if c.valueType != CellTypeEmpty {
		xmlCell.ValueType = c.valueType.odsName()

		switch c.valueType {
		case CellTypeFloat, CellTypeCurrency, CellTypePercentage:
			xmlCell.Value = c.rawValue
		case CellTypeDate:
			xmlCell.DateValue = c.rawValue
		case CellTypeBool:
			xmlCell.BooleanValue = c.rawValue
		case CellTypeString:
			xmlCell.StringValue = c.rawValue
		}

		displayValue := cellValueToString(c.valueType, c.rawValue)
		xmlCell.Paragraphs = []oxml.TextP{{Text: displayValue}}
	}

	if c.formula != "" {
		xmlCell.Formula = "of:=" + c.formula
	}

	if c.styleID > 0 {
		style := sm.get(c.styleID)
		if style != nil {
			styleName := fmt.Sprintf("ce%d", len(*autoStyles)+1)
			*autoStyles = append(*autoStyles, convertStyle(styleName, style))
			xmlCell.StyleName = styleName
		}
	}

	if c.colSpan > 1 {
		xmlCell.NumberColumnsSpanned = c.colSpan
	}
	if c.rowSpan > 1 {
		xmlCell.NumberRowsSpanned = c.rowSpan
	}

	return xmlCell
}

func convertStyle(name string, s *Style) oxml.Style {
	xs := oxml.Style{
		Name:   name,
		Family: "table-cell",
	}

	if s.Font != nil {
		xs.TextProperties = &oxml.TextProperties{
			FontName:   s.Font.Family,
			FontSize:   s.Font.Size,
			FontWeight: s.Font.Bold,
			FontStyle:  s.Font.Italic,
			Color:      s.Font.Color,
		}
		if s.Font.Underline {
			xs.TextProperties.TextUnderlineStyle = "solid"
		}
		if s.Font.Strikethrough {
			xs.TextProperties.TextLineThroughStyle = "solid"
		}
	}

	if s.Fill != nil {
		if xs.TableCellProperties == nil {
			xs.TableCellProperties = &oxml.TableCellProperties{}
		}
		xs.TableCellProperties.BackgroundColor = s.Fill.Color
	}

	if s.Border != nil {
		if xs.TableCellProperties == nil {
			xs.TableCellProperties = &oxml.TableCellProperties{}
		}
		border := formatBorder(s.Border)
		xs.TableCellProperties.Border = border
	}

	if s.Alignment != nil {
		xs.ParagraphProperties = &oxml.ParagraphProperties{
			TextAlign: s.Alignment.Horizontal,
		}
		if xs.TableCellProperties == nil {
			xs.TableCellProperties = &oxml.TableCellProperties{}
		}
		xs.TableCellProperties.VerticalAlign = s.Alignment.Vertical
		if s.Alignment.WrapText {
			xs.TableCellProperties.WrapOption = "wrap"
		}
	}

	return xs
}

func formatBorder(b *Border) string {
	if b.Style == "" {
		return ""
	}
	width := b.Width
	if width == "" {
		width = "0.06pt"
	}
	color := b.Color
	if color == "" {
		color = "#000000"
	}
	return fmt.Sprintf("%s %s %s", width, b.Style, color)
}

func (f *File) marshalContent() ([]byte, error) {
	doc, autoStyles := f.buildContentXML()
	var buf bytes.Buffer
	if err := oxml.WriteContentXML(&buf, doc, autoStyles); err != nil {
		return nil, err
	}
	return buf.Bytes(), nil
}

func (f *File) marshalMeta() ([]byte, error) {
	f.metadata.Meta.Date = time.Now().UTC().Format("2006-01-02T15:04:05")
	var buf bytes.Buffer
	if err := oxml.WriteMetaXML(&buf, f.metadata); err != nil {
		return nil, err
	}
	return buf.Bytes(), nil
}

func (f *File) marshalStyles() ([]byte, error) {
	var buf bytes.Buffer
	if err := oxml.WriteStylesXML(&buf, f.docStyles); err != nil {
		return nil, err
	}
	return buf.Bytes(), nil
}

func (f *File) marshalManifest() ([]byte, error) {
	m := oxml.DefaultManifest()
	var buf bytes.Buffer
	if err := oxml.WriteManifestXML(&buf, &m); err != nil {
		return nil, err
	}
	return buf.Bytes(), nil
}
