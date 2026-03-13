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
	sheets        []*sheet
	activeSheet   int
	styles        *styleManager
	metadata      *oxml.DocumentMeta
	docStyles     *oxml.DocumentStyles
	rawFiles      map[string][]byte
	path          string
	closed        bool
	autoRecalc    bool
	contentStyles map[string]oxml.Style
	namedRanges   []namedRange
	autoFilters   []autoFilter
}

type sheet struct {
	name        string
	columns     []column
	rows        map[int]*row
	merges      []mergeRange
	maxRow      int
	maxCol      int
	validations []*dataValidation
	freezeCol   int
	freezeRow   int
	printRange  *printRange
	pageSetup   *PageSetup
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
	valueType      CellType
	rawValue       string
	formula        string
	styleID        int
	colSpan        int
	rowSpan        int
	hyperlink      *Hyperlink
	numberFormat   string
	comment        *Comment
	styleName      string
	validationName string
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
		case "mimetype", "content.xml", "styles.xml", "meta.xml", "META-INF/manifest.xml", "settings.xml":
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

	if settingsData, ok := result.Files["settings.xml"]; ok {
		parseSettingsXML(f, settingsData)
	}

	return f, nil
}

type xmlParagraph struct {
	Text  string    `xml:",chardata"`
	Links []xmlLink `xml:",any"`
}

type xmlLink struct {
	XMLName xml.Name `xml:""`
	Href    string   `xml:"http://www.w3.org/1999/xlink href,attr"`
	Text    string   `xml:",chardata"`
}

type xmlAnnotation struct {
	Creator string         `xml:"creator"`
	Date    string         `xml:"date"`
	Paras   []xmlParagraph `xml:"p"`
}

type xmlTableCell struct {
	XMLName               xml.Name        `xml:"table-cell"`
	ValueType             string          `xml:"value-type,attr"`
	Value                 string          `xml:"value,attr"`
	DateValue             string          `xml:"date-value,attr"`
	BooleanValue          string          `xml:"boolean-value,attr"`
	StringValue           string          `xml:"string-value,attr"`
	Formula               string          `xml:"formula,attr"`
	StyleName             string          `xml:"style-name,attr"`
	ContentValidationName string          `xml:"content-validation-name,attr"`
	NumberColumnsRepeated int             `xml:"number-columns-repeated,attr"`
	NumberColumnsSpanned  int             `xml:"number-columns-spanned,attr"`
	NumberRowsSpanned     int             `xml:"number-rows-spanned,attr"`
	Annotations           []xmlAnnotation `xml:"annotation"`
	Paragraphs            []xmlParagraph  `xml:"p"`
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
	XMLName    xml.Name         `xml:"table"`
	Name       string           `xml:"name,attr"`
	PrintRange string           `xml:"print-ranges,attr"`
	Columns    []xmlTableColumn `xml:"table-column"`
	Rows       []xmlTableRow    `xml:"table-row"`
}

type xmlNamedRange struct {
	Name             string `xml:"name,attr"`
	BaseCellAddress  string `xml:"base-cell-address,attr"`
	CellRangeAddress string `xml:"cell-range-address,attr"`
}

type xmlContentValidation struct {
	Name         string       `xml:"name,attr"`
	Condition    string       `xml:"condition,attr"`
	AllowEmpty   string       `xml:"allow-empty-cell,attr"`
	BaseCellAddr string       `xml:"base-cell-address,attr"`
	ErrorMessage *xmlErrorMsg `xml:"error-message"`
	HelpMessage  *xmlHelpMsg  `xml:"help-message"`
}

type xmlErrorMsg struct {
	Display     string         `xml:"display,attr"`
	MessageType string         `xml:"message-type,attr"`
	Title       string         `xml:"title,attr"`
	Paras       []xmlParagraph `xml:"p"`
}

type xmlHelpMsg struct {
	Display string         `xml:"display,attr"`
	Title   string         `xml:"title,attr"`
	Paras   []xmlParagraph `xml:"p"`
}

type xmlDatabaseRange struct {
	Name                 string `xml:"name,attr"`
	TargetRangeAddress   string `xml:"target-range-address,attr"`
	DisplayFilterButtons string `xml:"display-filter-buttons,attr"`
}

type xmlAutoStyles struct {
	Styles []xmlStyleDef `xml:"style"`
}

type xmlStyleDef struct {
	XMLName         xml.Name `xml:"style"`
	Name            string   `xml:"name,attr"`
	Family          string   `xml:"family,attr"`
	ParentStyleName string   `xml:"parent-style-name,attr"`
	DataStyleName   string   `xml:"data-style-name,attr"`
}

type xmlContent struct {
	XMLName    xml.Name      `xml:"document-content"`
	AutoStyles xmlAutoStyles `xml:"automatic-styles"`
	Body       struct {
		Spreadsheet struct {
			Tables           []xmlTable `xml:"table"`
			NamedExpressions struct {
				NamedRanges []xmlNamedRange `xml:"named-range"`
			} `xml:"named-expressions"`
			ContentValidations struct {
				Validations []xmlContentValidation `xml:"content-validation"`
			} `xml:"content-validations"`
			DatabaseRanges struct {
				Ranges []xmlDatabaseRange `xml:"database-range"`
			} `xml:"database-ranges"`
		} `xml:"spreadsheet"`
	} `xml:"body"`
}

func parseContentXML(f *File, data []byte) error {
	var content xmlContent
	if err := xml.Unmarshal(data, &content); err != nil {
		return fmt.Errorf("unmarshal content: %w", err)
	}

	f.contentStyles = make(map[string]oxml.Style)
	for _, s := range content.AutoStyles.Styles {
		f.contentStyles[s.Name] = oxml.Style{
			Name:            s.Name,
			Family:          s.Family,
			ParentStyleName: s.ParentStyleName,
			DataStyleName:   s.DataStyleName,
		}
	}

	validationMap := make(map[string]*xmlContentValidation)
	for i := range content.Body.Spreadsheet.ContentValidations.Validations {
		v := &content.Body.Spreadsheet.ContentValidations.Validations[i]
		validationMap[v.Name] = v
	}

	for _, xmlTbl := range content.Body.Spreadsheet.Tables {
		s := &sheet{
			name:    xmlTbl.Name,
			rows:    make(map[int]*row),
			columns: make([]column, 0),
		}

		if xmlTbl.PrintRange != "" {
			_, sc, sr, ec, er, err := parseODSRangeAddress(xmlTbl.PrintRange)
			if err == nil {
				s.printRange = &printRange{
					startCol: sc, startRow: sr,
					endCol: ec, endRow: er,
				}
			}
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
							newCell := &cell{
								valueType:      c.valueType,
								rawValue:       c.rawValue,
								formula:        c.formula,
								styleID:        c.styleID,
								colSpan:        c.colSpan,
								rowSpan:        c.rowSpan,
								hyperlink:      c.hyperlink,
								comment:        c.comment,
								styleName:      c.styleName,
								validationName: c.validationName,
							}
							r.cells[colIdx+rep] = newCell
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

		for _, v := range content.Body.Spreadsheet.ContentValidations.Validations {
			if _, ok := validationMap[v.Name]; ok {
				dv := parseXMLValidation(&v)
				s.validations = append(s.validations, &dataValidation{
					name:       v.Name,
					validation: dv,
				})
			}
		}

		f.sheets = append(f.sheets, s)
	}

	for _, nr := range content.Body.Spreadsheet.NamedExpressions.NamedRanges {
		sheet, sc, sr, ec, er, err := parseODSRangeAddress(nr.CellRangeAddress)
		if err != nil {
			continue
		}
		f.namedRanges = append(f.namedRanges, namedRange{
			name: nr.Name, sheet: sheet,
			startCol: sc, startRow: sr,
			endCol: ec, endRow: er,
		})
	}

	for _, dr := range content.Body.Spreadsheet.DatabaseRanges.Ranges {
		if dr.DisplayFilterButtons != "true" {
			continue
		}
		sheet, sc, sr, ec, er, err := parseODSRangeAddress(dr.TargetRangeAddress)
		if err != nil {
			continue
		}
		f.autoFilters = append(f.autoFilters, autoFilter{
			sheet: sheet, startCol: sc, startRow: sr,
			endCol: ec, endRow: er,
		})
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

func parseXMLValidation(v *xmlContentValidation) *DataValidation {
	dv := &DataValidation{
		AllowEmpty: v.AllowEmpty == "true",
	}

	if v.Condition != "" {
		dv.Type, dv.Operator, dv.Formula1, dv.Formula2 = parseValidationCondition(v.Condition)
	}

	if v.ErrorMessage != nil {
		dv.ErrorStyle = v.ErrorMessage.MessageType
		dv.ErrorTitle = v.ErrorMessage.Title
		texts := make([]string, 0, len(v.ErrorMessage.Paras))
		for _, p := range v.ErrorMessage.Paras {
			texts = append(texts, p.Text)
		}
		dv.ErrorMessage = strings.Join(texts, "\n")
	}

	if v.HelpMessage != nil {
		dv.InputTitle = v.HelpMessage.Title
		texts := make([]string, 0, len(v.HelpMessage.Paras))
		for _, p := range v.HelpMessage.Paras {
			texts = append(texts, p.Text)
		}
		dv.InputMessage = strings.Join(texts, "\n")
	}

	return dv
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
	hasAnnotation := len(xc.Annotations) > 0
	hasValidation := xc.ContentValidationName != ""
	hasStyle := xc.StyleName != ""

	if xc.ValueType == "" && len(xc.Paragraphs) == 0 && xc.Formula == "" && !hasAnnotation && !hasValidation && !hasStyle {
		return nil
	}

	c := &cell{
		valueType:      cellTypeFromODS(xc.ValueType),
		colSpan:        xc.NumberColumnsSpanned,
		rowSpan:        xc.NumberRowsSpanned,
		styleName:      xc.StyleName,
		validationName: xc.ContentValidationName,
	}

	if hasAnnotation {
		ann := &xc.Annotations[0]
		texts := make([]string, 0, len(ann.Paras))
		for _, p := range ann.Paras {
			texts = append(texts, p.Text)
		}
		c.comment = &Comment{
			Author: ann.Creator,
			Date:   ann.Date,
			Text:   strings.Join(texts, "\n"),
		}
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
			if len(p.Links) > 0 {
				for _, link := range p.Links {
					if link.Href != "" {
						c.hyperlink = &Hyperlink{
							URL:     link.Href,
							Display: link.Text,
						}
						texts = append(texts, link.Text)
						break
					}
				}
			} else {
				texts = append(texts, p.Text)
			}
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

func parseSettingsXML(f *File, data []byte) {
	type configItem struct {
		Name  string `xml:"name,attr"`
		Type  string `xml:"type,attr"`
		Value string `xml:",chardata"`
	}
	type configItemMapEntry struct {
		Name  string       `xml:"name,attr"`
		Items []configItem `xml:"config-item"`
	}
	type configItemMapNamed struct {
		Name    string               `xml:"name,attr"`
		Entries []configItemMapEntry `xml:"config-item-map-entry"`
	}
	type configItemSet struct {
		Name  string               `xml:"name,attr"`
		Items []configItem         `xml:"config-item"`
		Maps  []configItemMapNamed `xml:"config-item-map-named"`
	}
	type documentSettings struct {
		XMLName xml.Name        `xml:"document-settings"`
		Sets    []configItemSet `xml:"settings>config-item-set"`
	}

	var settings documentSettings
	if err := xml.Unmarshal(data, &settings); err != nil {
		return
	}

	for _, set := range settings.Sets {
		if set.Name != "ooo:view-settings" {
			continue
		}
		for _, m := range set.Maps {
			if m.Name != "Views" {
				continue
			}
			for _, viewEntry := range m.Entries {
				for _, tableMap := range viewEntry.Items {
					_ = tableMap
				}
			}
		}
	}
}

func (f *File) marshalSettings() ([]byte, error) {
	hasFreezeSettings := false
	for _, s := range f.sheets {
		if s.freezeCol > 0 || s.freezeRow > 0 {
			hasFreezeSettings = true
			break
		}
	}
	if !hasFreezeSettings {
		return nil, nil
	}

	var buf bytes.Buffer
	buf.WriteString(xml.Header)
	xw := oxml.NewWriter(&buf)

	rootAttrs := []xml.Attr{
		oxml.NSAttr("office"),
		oxml.NSAttr("config"),
		oxml.Attr("office", "version", "1.2"),
	}

	if err := xw.StartElement("office", "document-settings", rootAttrs...); err != nil {
		return nil, err
	}
	if err := xw.StartElement("office", "settings"); err != nil {
		return nil, err
	}

	if err := xw.StartElement("config", "config-item-set",
		oxml.Attr("config", "name", "ooo:view-settings"),
	); err != nil {
		return nil, err
	}

	if err := xw.StartElement("config", "config-item-map-named",
		oxml.Attr("config", "name", "Views"),
	); err != nil {
		return nil, err
	}

	if err := xw.StartElement("config", "config-item-map-entry"); err != nil {
		return nil, err
	}

	if err := writeConfigItem(xw, "ViewId", "string", "view1"); err != nil {
		return nil, err
	}

	if err := xw.StartElement("config", "config-item-map-named",
		oxml.Attr("config", "name", "Tables"),
	); err != nil {
		return nil, err
	}

	for _, s := range f.sheets {
		if s.freezeCol == 0 && s.freezeRow == 0 {
			continue
		}

		if err := xw.StartElement("config", "config-item-map-entry",
			oxml.Attr("config", "name", s.name),
		); err != nil {
			return nil, err
		}

		if s.freezeCol > 0 {
			if err := writeConfigItem(xw, "HorizontalSplitMode", "short", "2"); err != nil {
				return nil, err
			}
			if err := writeConfigItem(xw, "HorizontalSplitPosition", "int", fmt.Sprintf("%d", s.freezeCol)); err != nil {
				return nil, err
			}
		}
		if s.freezeRow > 0 {
			if err := writeConfigItem(xw, "VerticalSplitMode", "short", "2"); err != nil {
				return nil, err
			}
			if err := writeConfigItem(xw, "VerticalSplitPosition", "int", fmt.Sprintf("%d", s.freezeRow)); err != nil {
				return nil, err
			}
		}

		posRight := 0
		if s.freezeCol > 0 {
			posRight = 1
		}
		posBottom := 0
		if s.freezeRow > 0 {
			posBottom = 1
		}
		activeSplit := posRight | (posBottom << 1)
		if activeSplit == 0 {
			activeSplit = 2
		}
		if err := writeConfigItem(xw, "PositionRight", "int", fmt.Sprintf("%d", s.freezeCol)); err != nil {
			return nil, err
		}
		if err := writeConfigItem(xw, "PositionBottom", "int", fmt.Sprintf("%d", s.freezeRow)); err != nil {
			return nil, err
		}

		_ = activeSplit

		if err := xw.EndElement("config", "config-item-map-entry"); err != nil {
			return nil, err
		}
	}

	if err := xw.EndElement("config", "config-item-map-named"); err != nil {
		return nil, err
	}
	if err := xw.EndElement("config", "config-item-map-entry"); err != nil {
		return nil, err
	}
	if err := xw.EndElement("config", "config-item-map-named"); err != nil {
		return nil, err
	}
	if err := xw.EndElement("config", "config-item-set"); err != nil {
		return nil, err
	}
	if err := xw.EndElement("office", "settings"); err != nil {
		return nil, err
	}
	if err := xw.EndElement("office", "document-settings"); err != nil {
		return nil, err
	}

	if err := xw.Flush(); err != nil {
		return nil, err
	}

	return buf.Bytes(), nil
}

func writeConfigItem(xw *oxml.Writer, name, typ, value string) error {
	if err := xw.StartElement("config", "config-item",
		oxml.Attr("config", "name", name),
		oxml.Attr("config", "type", typ),
	); err != nil {
		return err
	}
	if err := xw.CharData(value); err != nil {
		return err
	}
	return xw.EndElement("config", "config-item")
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

func (f *File) buildContentXML() (*oxml.DocumentContent, []oxml.Style, []oxml.ContentValidation, []oxml.NamedRange, []oxml.DatabaseRange) {
	doc := &oxml.DocumentContent{}
	var autoStyles []oxml.Style
	var allValidations []oxml.ContentValidation
	var allNamedRanges []oxml.NamedRange
	var allDatabaseRanges []oxml.DatabaseRange

	for _, nr := range f.namedRanges {
		allNamedRanges = append(allNamedRanges, oxml.NamedRange{
			Name:             nr.name,
			BaseCellAddress:  formatODSCellAddress(nr.sheet, nr.startCol, nr.startRow),
			CellRangeAddress: formatODSRangeAddress(nr.sheet, nr.startCol, nr.startRow, nr.endCol, nr.endRow),
		})
	}

	for i, af := range f.autoFilters {
		allDatabaseRanges = append(allDatabaseRanges, oxml.DatabaseRange{
			Name:                 fmt.Sprintf("__Anonymous_Sheet_DB_%d", i),
			TargetRangeAddress:   formatODSRangeAddress(af.sheet, af.startCol, af.startRow, af.endCol, af.endRow),
			DisplayFilterButtons: "true",
		})
	}

	for _, s := range f.sheets {
		table := oxml.Table{
			Name: s.name,
		}

		if s.printRange != nil {
			table.PrintRanges = formatODSRangeAddress(s.name, s.printRange.startCol, s.printRange.startRow, s.printRange.endCol, s.printRange.endRow)
		}

		for _, dv := range s.validations {
			cv := oxml.ContentValidation{
				Name:       dv.name,
				Condition:  buildValidationCondition(dv.validation),
				AllowEmpty: dv.validation.AllowEmpty,
			}
			if dv.validation.ErrorMessage != "" || dv.validation.ErrorTitle != "" {
				cv.ErrorMessage = &oxml.ErrorMessage{
					Display:     true,
					MessageType: dv.validation.ErrorStyle,
					Title:       dv.validation.ErrorTitle,
					Text:        dv.validation.ErrorMessage,
				}
			}
			if dv.validation.InputMessage != "" || dv.validation.InputTitle != "" {
				cv.HelpMessage = &oxml.HelpMessage{
					Display: true,
					Title:   dv.validation.InputTitle,
					Text:    dv.validation.InputMessage,
				}
			}
			allValidations = append(allValidations, cv)
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

	for name, cs := range f.contentStyles {
		found := false
		for _, as := range autoStyles {
			if as.Name == name {
				found = true
				break
			}
		}
		if !found {
			autoStyles = append(autoStyles, cs)
		}
	}

	return doc, autoStyles, allValidations, allNamedRanges, allDatabaseRanges
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
	} else if c.styleName != "" {
		xmlCell.StyleName = c.styleName
	}

	if c.validationName != "" {
		xmlCell.ContentValidationName = c.validationName
	}

	if c.hyperlink != nil {
		xmlCell.Paragraphs = []oxml.TextP{{
			Link: &oxml.TextA{
				Href: c.hyperlink.URL,
				Type: "simple",
				Text: c.hyperlink.Display,
			},
		}}
	}

	if c.comment != nil {
		xmlCell.Annotation = &oxml.Annotation{
			Creator: c.comment.Author,
			Date:    c.comment.Date,
			Text:    c.comment.Text,
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
	doc, autoStyles, validations, namedRanges, databaseRanges := f.buildContentXML()
	var buf bytes.Buffer
	if err := oxml.WriteContentXML(&buf, doc, autoStyles, validations, namedRanges, databaseRanges); err != nil {
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
