package xml

import (
	"encoding/xml"
	"fmt"
	"io"
	"strings"
)

type Writer struct {
	w       io.Writer
	encoder *xml.Encoder
}

func NewWriter(w io.Writer) *Writer {
	enc := xml.NewEncoder(w)
	enc.Indent("", " ")
	return &Writer{w: w, encoder: enc}
}

func (xw *Writer) WriteRaw(s string) error {
	_, err := io.WriteString(xw.w, s)
	return err
}

func (xw *Writer) StartElement(prefix, local string, attrs ...xml.Attr) error {
	name := local
	if prefix != "" {
		name = prefix + ":" + local
	}
	return xw.encoder.EncodeToken(xml.StartElement{
		Name: xml.Name{Local: name},
		Attr: attrs,
	})
}

func (xw *Writer) EndElement(prefix, local string) error {
	name := local
	if prefix != "" {
		name = prefix + ":" + local
	}
	return xw.encoder.EncodeToken(xml.EndElement{
		Name: xml.Name{Local: name},
	})
}

func (xw *Writer) CharData(text string) error {
	return xw.encoder.EncodeToken(xml.CharData(text))
}

func (xw *Writer) Flush() error {
	return xw.encoder.Flush()
}

func Attr(prefix, local, value string) xml.Attr {
	name := local
	if prefix != "" {
		name = prefix + ":" + local
	}
	return xml.Attr{Name: xml.Name{Local: name}, Value: value}
}

func NSAttr(prefix string) xml.Attr {
	ns, ok := PrefixToNS[prefix]
	if !ok {
		return xml.Attr{}
	}
	return xml.Attr{
		Name:  xml.Name{Local: "xmlns:" + prefix},
		Value: ns,
	}
}

func WriteContentXML(w io.Writer, doc *DocumentContent, autoStyles []Style) error {
	if _, err := io.WriteString(w, xml.Header); err != nil {
		return err
	}

	xw := NewWriter(w)

	rootAttrs := []xml.Attr{
		NSAttr("office"), NSAttr("style"), NSAttr("text"),
		NSAttr("table"), NSAttr("draw"), NSAttr("fo"),
		NSAttr("xlink"), NSAttr("dc"), NSAttr("meta"),
		NSAttr("number"), NSAttr("svg"), NSAttr("chart"),
		NSAttr("dr3d"), NSAttr("form"), NSAttr("script"),
		Attr("office", "version", "1.2"),
	}

	if err := xw.StartElement("office", "document-content", rootAttrs...); err != nil {
		return err
	}

	if err := writeAutoStyles(xw, autoStyles); err != nil {
		return err
	}

	if err := xw.StartElement("office", "body"); err != nil {
		return err
	}
	if err := xw.StartElement("office", "spreadsheet"); err != nil {
		return err
	}

	for _, table := range doc.Body.Spreadsheet.Tables {
		if err := writeTable(xw, &table); err != nil {
			return err
		}
	}

	if err := xw.EndElement("office", "spreadsheet"); err != nil {
		return err
	}
	if err := xw.EndElement("office", "body"); err != nil {
		return err
	}
	if err := xw.EndElement("office", "document-content"); err != nil {
		return err
	}

	return xw.Flush()
}

func writeAutoStyles(xw *Writer, styles []Style) error {
	if err := xw.StartElement("office", "automatic-styles"); err != nil {
		return err
	}

	for _, s := range styles {
		if err := writeStyle(xw, &s); err != nil {
			return err
		}
	}

	return xw.EndElement("office", "automatic-styles")
}

func writeStyle(xw *Writer, s *Style) error {
	attrs := []xml.Attr{
		Attr("style", "name", s.Name),
		Attr("style", "family", s.Family),
	}
	if s.ParentStyleName != "" {
		attrs = append(attrs, Attr("style", "parent-style-name", s.ParentStyleName))
	}
	if s.DataStyleName != "" {
		attrs = append(attrs, Attr("style", "data-style-name", s.DataStyleName))
	}

	if err := xw.StartElement("style", "style", attrs...); err != nil {
		return err
	}

	if s.TableColumnProperties != nil {
		colAttrs := []xml.Attr{}
		if s.TableColumnProperties.ColumnWidth != "" {
			colAttrs = append(colAttrs, Attr("style", "column-width", s.TableColumnProperties.ColumnWidth))
		}
		if err := xw.StartElement("style", "table-column-properties", colAttrs...); err != nil {
			return err
		}
		if err := xw.EndElement("style", "table-column-properties"); err != nil {
			return err
		}
	}

	if s.TableRowProperties != nil {
		rowAttrs := []xml.Attr{}
		if s.TableRowProperties.RowHeight != "" {
			rowAttrs = append(rowAttrs, Attr("style", "row-height", s.TableRowProperties.RowHeight))
		}
		if s.TableRowProperties.UseOptimalHeight != "" {
			rowAttrs = append(rowAttrs, Attr("style", "use-optimal-row-height", s.TableRowProperties.UseOptimalHeight))
		}
		if err := xw.StartElement("style", "table-row-properties", rowAttrs...); err != nil {
			return err
		}
		if err := xw.EndElement("style", "table-row-properties"); err != nil {
			return err
		}
	}

	if s.TableCellProperties != nil {
		if err := writeCellProperties(xw, s.TableCellProperties); err != nil {
			return err
		}
	}

	if s.TextProperties != nil {
		if err := writeTextProperties(xw, s.TextProperties); err != nil {
			return err
		}
	}

	if s.ParagraphProperties != nil {
		pAttrs := []xml.Attr{}
		if s.ParagraphProperties.TextAlign != "" {
			pAttrs = append(pAttrs, Attr("fo", "text-align", s.ParagraphProperties.TextAlign))
		}
		if err := xw.StartElement("style", "paragraph-properties", pAttrs...); err != nil {
			return err
		}
		if err := xw.EndElement("style", "paragraph-properties"); err != nil {
			return err
		}
	}

	return xw.EndElement("style", "style")
}

func writeCellProperties(xw *Writer, cp *TableCellProperties) error {
	attrs := []xml.Attr{}
	if cp.BackgroundColor != "" {
		attrs = append(attrs, Attr("fo", "background-color", cp.BackgroundColor))
	}
	if cp.Border != "" {
		attrs = append(attrs, Attr("fo", "border", cp.Border))
	}
	if cp.BorderTop != "" {
		attrs = append(attrs, Attr("fo", "border-top", cp.BorderTop))
	}
	if cp.BorderBottom != "" {
		attrs = append(attrs, Attr("fo", "border-bottom", cp.BorderBottom))
	}
	if cp.BorderLeft != "" {
		attrs = append(attrs, Attr("fo", "border-left", cp.BorderLeft))
	}
	if cp.BorderRight != "" {
		attrs = append(attrs, Attr("fo", "border-right", cp.BorderRight))
	}
	if cp.VerticalAlign != "" {
		attrs = append(attrs, Attr("style", "vertical-align", cp.VerticalAlign))
	}
	if cp.WrapOption != "" {
		attrs = append(attrs, Attr("fo", "wrap-option", cp.WrapOption))
	}

	if err := xw.StartElement("style", "table-cell-properties", attrs...); err != nil {
		return err
	}
	return xw.EndElement("style", "table-cell-properties")
}

func writeTextProperties(xw *Writer, tp *TextProperties) error {
	attrs := []xml.Attr{}
	if tp.FontName != "" {
		attrs = append(attrs, Attr("style", "font-name", tp.FontName))
	}
	if tp.FontSize != "" {
		attrs = append(attrs, Attr("fo", "font-size", tp.FontSize))
	}
	if tp.FontWeight != "" {
		attrs = append(attrs, Attr("fo", "font-weight", tp.FontWeight))
	}
	if tp.FontStyle != "" {
		attrs = append(attrs, Attr("fo", "font-style", tp.FontStyle))
	}
	if tp.Color != "" {
		attrs = append(attrs, Attr("fo", "color", tp.Color))
	}
	if tp.TextUnderlineStyle != "" {
		attrs = append(attrs, Attr("style", "text-underline-style", tp.TextUnderlineStyle))
	}
	if tp.TextLineThroughStyle != "" {
		attrs = append(attrs, Attr("style", "text-line-through-style", tp.TextLineThroughStyle))
	}

	if err := xw.StartElement("style", "text-properties", attrs...); err != nil {
		return err
	}
	return xw.EndElement("style", "text-properties")
}

func writeTable(xw *Writer, table *Table) error {
	if err := xw.StartElement("table", "table",
		Attr("table", "name", table.Name),
	); err != nil {
		return err
	}

	for _, col := range table.Columns {
		attrs := []xml.Attr{}
		if col.StyleName != "" {
			attrs = append(attrs, Attr("table", "style-name", col.StyleName))
		}
		if col.DefaultCellStyleName != "" {
			attrs = append(attrs, Attr("table", "default-cell-style-name", col.DefaultCellStyleName))
		}
		if col.NumberColumnsRepeated > 1 {
			attrs = append(attrs, Attr("table", "number-columns-repeated", fmt.Sprintf("%d", col.NumberColumnsRepeated)))
		}
		if err := xw.StartElement("table", "table-column", attrs...); err != nil {
			return err
		}
		if err := xw.EndElement("table", "table-column"); err != nil {
			return err
		}
	}

	for _, row := range table.Rows {
		if err := writeRow(xw, &row); err != nil {
			return err
		}
	}

	return xw.EndElement("table", "table")
}

func writeRow(xw *Writer, row *TableRow) error {
	attrs := []xml.Attr{}
	if row.StyleName != "" {
		attrs = append(attrs, Attr("table", "style-name", row.StyleName))
	}
	if row.NumberRowsRepeated > 1 {
		attrs = append(attrs, Attr("table", "number-rows-repeated", fmt.Sprintf("%d", row.NumberRowsRepeated)))
	}

	if err := xw.StartElement("table", "table-row", attrs...); err != nil {
		return err
	}

	for _, cell := range row.Cells {
		if err := writeCell(xw, &cell); err != nil {
			return err
		}
	}

	return xw.EndElement("table", "table-row")
}

func writeCell(xw *Writer, cell *TableCell) error {
	attrs := []xml.Attr{}

	if cell.ValueType != "" {
		attrs = append(attrs, Attr("office", "value-type", cell.ValueType))

		switch cell.ValueType {
		case "float", "currency", "percentage":
			if cell.Value != "" {
				attrs = append(attrs, Attr("office", "value", cell.Value))
			}
		case "date":
			if cell.DateValue != "" {
				attrs = append(attrs, Attr("office", "date-value", cell.DateValue))
			}
		case "boolean":
			if cell.BooleanValue != "" {
				attrs = append(attrs, Attr("office", "boolean-value", cell.BooleanValue))
			}
		case "string":
			if cell.StringValue != "" {
				attrs = append(attrs, Attr("office", "string-value", cell.StringValue))
			}
		}
	}

	if cell.Formula != "" {
		formula := cell.Formula
		if !strings.HasPrefix(formula, "of:") {
			formula = "of:=" + formula
		}
		attrs = append(attrs, Attr("table", "formula", formula))
	}

	if cell.StyleName != "" {
		attrs = append(attrs, Attr("table", "style-name", cell.StyleName))
	}
	if cell.NumberColumnsRepeated > 1 {
		attrs = append(attrs, Attr("table", "number-columns-repeated", fmt.Sprintf("%d", cell.NumberColumnsRepeated)))
	}
	if cell.NumberColumnsSpanned > 1 {
		attrs = append(attrs, Attr("table", "number-columns-spanned", fmt.Sprintf("%d", cell.NumberColumnsSpanned)))
	}
	if cell.NumberRowsSpanned > 1 {
		attrs = append(attrs, Attr("table", "number-rows-spanned", fmt.Sprintf("%d", cell.NumberRowsSpanned)))
	}

	if err := xw.StartElement("table", "table-cell", attrs...); err != nil {
		return err
	}

	for _, p := range cell.Paragraphs {
		if err := xw.StartElement("text", "p"); err != nil {
			return err
		}
		if p.Link != nil {
			linkAttrs := []xml.Attr{
				Attr("xlink", "href", p.Link.Href),
				Attr("xlink", "type", "simple"),
			}
			if err := xw.StartElement("text", "a", linkAttrs...); err != nil {
				return err
			}
			if p.Link.Text != "" {
				if err := xw.CharData(p.Link.Text); err != nil {
					return err
				}
			}
			if err := xw.EndElement("text", "a"); err != nil {
				return err
			}
		} else if p.Text != "" {
			if err := xw.CharData(p.Text); err != nil {
				return err
			}
		}
		if err := xw.EndElement("text", "p"); err != nil {
			return err
		}
	}

	return xw.EndElement("table", "table-cell")
}

func WriteMetaXML(w io.Writer, meta *DocumentMeta) error {
	if _, err := io.WriteString(w, xml.Header); err != nil {
		return err
	}

	xw := NewWriter(w)

	rootAttrs := []xml.Attr{
		NSAttr("office"), NSAttr("dc"), NSAttr("meta"),
		Attr("office", "version", "1.2"),
	}

	if err := xw.StartElement("office", "document-meta", rootAttrs...); err != nil {
		return err
	}
	if err := xw.StartElement("office", "meta"); err != nil {
		return err
	}

	if meta.Meta.Generator != "" {
		if err := writeSimpleElement(xw, "meta", "generator", meta.Meta.Generator); err != nil {
			return err
		}
	}
	if meta.Meta.Title != "" {
		if err := writeSimpleElement(xw, "dc", "title", meta.Meta.Title); err != nil {
			return err
		}
	}
	if meta.Meta.Description != "" {
		if err := writeSimpleElement(xw, "dc", "description", meta.Meta.Description); err != nil {
			return err
		}
	}
	if meta.Meta.Subject != "" {
		if err := writeSimpleElement(xw, "dc", "subject", meta.Meta.Subject); err != nil {
			return err
		}
	}
	if meta.Meta.Creator != "" {
		if err := writeSimpleElement(xw, "dc", "creator", meta.Meta.Creator); err != nil {
			return err
		}
	}
	if meta.Meta.CreationDate != "" {
		if err := writeSimpleElement(xw, "meta", "creation-date", meta.Meta.CreationDate); err != nil {
			return err
		}
	}
	if meta.Meta.Date != "" {
		if err := writeSimpleElement(xw, "dc", "date", meta.Meta.Date); err != nil {
			return err
		}
	}

	if err := xw.EndElement("office", "meta"); err != nil {
		return err
	}
	if err := xw.EndElement("office", "document-meta"); err != nil {
		return err
	}

	return xw.Flush()
}

func WriteStylesXML(w io.Writer, styles *DocumentStyles) error {
	if _, err := io.WriteString(w, xml.Header); err != nil {
		return err
	}

	xw := NewWriter(w)

	rootAttrs := []xml.Attr{
		NSAttr("office"), NSAttr("style"), NSAttr("text"),
		NSAttr("table"), NSAttr("draw"), NSAttr("fo"),
		NSAttr("xlink"), NSAttr("dc"), NSAttr("meta"),
		NSAttr("number"), NSAttr("svg"),
		Attr("office", "version", "1.2"),
	}

	if err := xw.StartElement("office", "document-styles", rootAttrs...); err != nil {
		return err
	}

	if err := xw.StartElement("office", "styles"); err != nil {
		return err
	}
	for _, ds := range styles.Styles.DefaultStyles {
		attrs := []xml.Attr{Attr("style", "family", ds.Family)}
		if err := xw.StartElement("style", "default-style", attrs...); err != nil {
			return err
		}
		if ds.ParagraphProperties != nil {
			pAttrs := []xml.Attr{}
			if ds.ParagraphProperties.TextAlign != "" {
				pAttrs = append(pAttrs, Attr("fo", "text-align", ds.ParagraphProperties.TextAlign))
			}
			if err := xw.StartElement("style", "paragraph-properties", pAttrs...); err != nil {
				return err
			}
			if err := xw.EndElement("style", "paragraph-properties"); err != nil {
				return err
			}
		}
		if ds.TextProperties != nil {
			if err := writeTextProperties(xw, ds.TextProperties); err != nil {
				return err
			}
		}
		if err := xw.EndElement("style", "default-style"); err != nil {
			return err
		}
	}
	if err := xw.EndElement("office", "styles"); err != nil {
		return err
	}

	if err := xw.StartElement("office", "automatic-styles"); err != nil {
		return err
	}
	if err := xw.EndElement("office", "automatic-styles"); err != nil {
		return err
	}

	if err := xw.StartElement("office", "master-styles"); err != nil {
		return err
	}
	for _, mp := range styles.MasterStyles.MasterPages {
		mpAttrs := []xml.Attr{
			Attr("style", "name", mp.Name),
		}
		if mp.PageLayoutName != "" {
			mpAttrs = append(mpAttrs, Attr("style", "page-layout-name", mp.PageLayoutName))
		}
		if err := xw.StartElement("style", "master-page", mpAttrs...); err != nil {
			return err
		}
		if err := xw.EndElement("style", "master-page"); err != nil {
			return err
		}
	}
	if err := xw.EndElement("office", "master-styles"); err != nil {
		return err
	}

	if err := xw.EndElement("office", "document-styles"); err != nil {
		return err
	}

	return xw.Flush()
}

func WriteManifestXML(w io.Writer, manifest *Manifest) error {
	if _, err := io.WriteString(w, xml.Header); err != nil {
		return err
	}

	xw := NewWriter(w)

	if err := xw.StartElement("manifest", "manifest",
		NSAttr("manifest"),
		Attr("manifest", "version", "1.2"),
	); err != nil {
		return err
	}

	for _, fe := range manifest.FileEntries {
		attrs := []xml.Attr{
			Attr("manifest", "full-path", fe.FullPath),
			Attr("manifest", "media-type", fe.MediaType),
		}
		if fe.Version != "" {
			attrs = append(attrs, Attr("manifest", "version", fe.Version))
		}
		if err := xw.StartElement("manifest", "file-entry", attrs...); err != nil {
			return err
		}
		if err := xw.EndElement("manifest", "file-entry"); err != nil {
			return err
		}
	}

	if err := xw.EndElement("manifest", "manifest"); err != nil {
		return err
	}

	return xw.Flush()
}

func writeSimpleElement(xw *Writer, prefix, local, text string) error {
	if err := xw.StartElement(prefix, local); err != nil {
		return err
	}
	if err := xw.CharData(text); err != nil {
		return err
	}
	return xw.EndElement(prefix, local)
}
