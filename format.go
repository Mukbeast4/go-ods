package goods

import (
	"fmt"
	"math"
	"strconv"
	"strings"
)

func (f *File) SetCellNumberFormat(sheet, cellRef, format string) error {
	if f.closed {
		return ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
		return ErrSheetNotFound
	}

	col, row, err := CellNameToCoordinates(cellRef)
	if err != nil {
		return err
	}

	r := s.getOrCreateRow(row)
	c, ok := r.cells[col]
	if !ok {
		c = &cell{}
		r.cells[col] = c
	}

	c.numberFormat = format

	if col > s.maxCol {
		s.maxCol = col
	}
	if row > s.maxRow {
		s.maxRow = row
	}

	return nil
}

func (f *File) GetCellNumberFormat(sheet, cellRef string) (string, error) {
	if f.closed {
		return "", ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
		return "", ErrSheetNotFound
	}

	col, row, err := CellNameToCoordinates(cellRef)
	if err != nil {
		return "", err
	}

	c := s.getCell(col, row)
	if c == nil {
		return "", nil
	}

	return c.numberFormat, nil
}

func formatNumber(value float64, format string) string {
	switch {
	case format == "":
		return strconv.FormatFloat(value, 'f', -1, 64)

	case format == "0":
		return strconv.FormatFloat(math.Round(value), 'f', 0, 64)

	case format == "0.00":
		return strconv.FormatFloat(value, 'f', 2, 64)

	case format == "#,##0":
		return formatWithThousands(value, 0)

	case format == "#,##0.00":
		return formatWithThousands(value, 2)

	case format == "0%":
		return strconv.FormatFloat(value*100, 'f', 0, 64) + "%"

	case format == "0.00%":
		return strconv.FormatFloat(value*100, 'f', 2, 64) + "%"

	case strings.HasPrefix(format, "$"):
		inner := strings.TrimPrefix(format, "$")
		return "$" + formatNumberInner(value, inner)

	case strings.Contains(format, "€") || strings.HasPrefix(format, "€"):
		inner := strings.Replace(format, "€", "", 1)
		return "€" + formatNumberInner(value, inner)

	case format == "YYYY-MM-DD" || format == "DD/MM/YYYY" || format == "HH:MM:SS":
		return strconv.FormatFloat(value, 'f', -1, 64)

	default:
		return strconv.FormatFloat(value, 'f', -1, 64)
	}
}

func formatNumberInner(value float64, format string) string {
	switch format {
	case "#,##0.00":
		return formatWithThousands(value, 2)
	case "#,##0":
		return formatWithThousands(value, 0)
	default:
		return strconv.FormatFloat(value, 'f', 2, 64)
	}
}

func formatWithThousands(value float64, decimals int) string {
	negative := value < 0
	if negative {
		value = -value
	}

	s := strconv.FormatFloat(value, 'f', decimals, 64)

	parts := strings.SplitN(s, ".", 2)
	intPart := parts[0]

	var result strings.Builder
	for i, c := range intPart {
		if i > 0 && (len(intPart)-i)%3 == 0 {
			result.WriteRune(',')
		}
		result.WriteRune(c)
	}

	formatted := result.String()
	if len(parts) > 1 {
		formatted += "." + parts[1]
	}

	if negative {
		return "-" + formatted
	}
	return formatted
}

func (f *File) GetCellFormattedValue(sheet, cellRef string) (string, error) {
	if f.closed {
		return "", ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
		return "", ErrSheetNotFound
	}

	col, row, err := CellNameToCoordinates(cellRef)
	if err != nil {
		return "", err
	}

	c := s.getCell(col, row)
	if c == nil {
		return "", nil
	}

	if c.numberFormat == "" {
		return cellValueToString(c.valueType, c.rawValue), nil
	}

	if c.valueType == CellTypeFloat || c.valueType == CellTypeCurrency || c.valueType == CellTypePercentage {
		v, parseErr := strconv.ParseFloat(c.rawValue, 64)
		if parseErr != nil {
			return c.rawValue, nil
		}
		return formatNumber(v, c.numberFormat), nil
	}

	return cellValueToString(c.valueType, c.rawValue), nil
}

func numberFormatToODSStyleName(format string) string {
	return fmt.Sprintf("nf_%s", strings.NewReplacer(
		"#", "H",
		",", "C",
		".", "D",
		"0", "Z",
		"%", "P",
		"$", "S",
		"€", "E",
		"-", "M",
		"/", "L",
		":", "T",
		" ", "_",
	).Replace(format))
}
