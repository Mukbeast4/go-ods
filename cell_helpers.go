package goods

import (
	"fmt"
	"strconv"
	"strings"
	"time"
)

// SetCellCurrency writes value as a currency cell formatted with the given
// symbol. When symbol is "", "$" is used. The cell value is the raw number
// (e.g. 1234.5), the symbol and thousands grouping are applied on display
// by GetCellFormattedValue.
func (f *File) SetCellCurrency(sheet, cellRef string, value float64, symbol string) error {
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

	if symbol == "" {
		symbol = "$"
	}

	rawValue := strconv.FormatFloat(value, 'f', -1, 64)
	s.setCellValue(col, row, CellTypeCurrency, rawValue)

	r := s.getOrCreateRow(row)
	r.cells[col].numberFormat = symbol + "#,##0.00"

	f.triggerRecalc(sheet)
	return nil
}

// SetCellPercentage writes value as a percentage cell. The stored value is a
// proportion: 0.25 renders as "25%". decimals sets the number of fractional
// digits shown (0 for "0%", 2 for "0.00%"); other values fall back to "0.00%".
func (f *File) SetCellPercentage(sheet, cellRef string, value float64, decimals int) error {
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

	format := "0%"
	if decimals > 0 {
		format = "0." + strings.Repeat("0", decimals) + "%"
	}

	rawValue := strconv.FormatFloat(value, 'f', -1, 64)
	s.setCellValue(col, row, CellTypePercentage, rawValue)

	r := s.getOrCreateRow(row)
	r.cells[col].numberFormat = format

	f.triggerRecalc(sheet)
	return nil
}

// SetCellDateTime writes t as a date cell with a custom display format. format
// is an ODS-compatible code such as "DD/MM/YYYY HH:MM:SS". Pass "" to use the
// default ISO encoding, equivalent to SetCellDate.
func (f *File) SetCellDateTime(sheet, cellRef string, t time.Time, format string) error {
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

	rawValue := t.Format("2006-01-02T15:04:05")
	s.setCellValue(col, row, CellTypeDate, rawValue)

	if format != "" {
		r := s.getOrCreateRow(row)
		r.cells[col].numberFormat = format
	}

	f.triggerRecalc(sheet)
	return nil
}

// SetCellDuration writes d as a time cell. The value is stored as an ISO 8601
// duration ("PT1H30M45S") which is the representation expected by ODS for
// office:time-value. Sub-second precision is truncated.
func (f *File) SetCellDuration(sheet, cellRef string, d time.Duration) error {
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

	s.setCellValue(col, row, CellTypeTime, formatISODuration(d))

	r := s.getOrCreateRow(row)
	r.cells[col].numberFormat = "[HH]:MM:SS"

	f.triggerRecalc(sheet)
	return nil
}

func formatISODuration(d time.Duration) string {
	if d < 0 {
		return "-" + formatISODuration(-d)
	}
	hours := int64(d / time.Hour)
	d -= time.Duration(hours) * time.Hour
	minutes := int64(d / time.Minute)
	d -= time.Duration(minutes) * time.Minute
	seconds := int64(d / time.Second)

	return fmt.Sprintf("PT%dH%dM%dS", hours, minutes, seconds)
}
