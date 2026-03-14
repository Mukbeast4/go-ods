package goods

import (
	"fmt"
	"time"
)

func (f *File) SetCellValue(sheet, cellRef string, value interface{}) error {
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

	valueType, rawValue := detectCellType(value)
	s.setCellValue(col, row, valueType, rawValue)
	f.triggerRecalc(sheet)
	return nil
}

func (f *File) GetCellValue(sheet, cellRef string) (string, error) {
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

	return cellValueToString(c.valueType, c.rawValue), nil
}

func (f *File) GetCellType(sheet, cellRef string) (CellType, error) {
	if f.closed {
		return CellTypeEmpty, ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
		return CellTypeEmpty, ErrSheetNotFound
	}

	col, row, err := CellNameToCoordinates(cellRef)
	if err != nil {
		return CellTypeEmpty, err
	}

	c := s.getCell(col, row)
	if c == nil {
		return CellTypeEmpty, nil
	}

	return c.valueType, nil
}

func (f *File) SetCellStr(sheet, cellRef, value string) error {
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

	s.setCellValue(col, row, CellTypeString, value)
	f.triggerRecalc(sheet)
	return nil
}

func (f *File) SetCellInt(sheet, cellRef string, value int64) error {
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

	_, rawValue := detectCellType(value)
	s.setCellValue(col, row, CellTypeFloat, rawValue)
	f.triggerRecalc(sheet)
	return nil
}

func (f *File) SetCellFloat(sheet, cellRef string, value float64) error {
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

	_, rawValue := detectCellType(value)
	s.setCellValue(col, row, CellTypeFloat, rawValue)
	f.triggerRecalc(sheet)
	return nil
}

func (f *File) SetCellBool(sheet, cellRef string, value bool) error {
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

	_, rawValue := detectCellType(value)
	s.setCellValue(col, row, CellTypeBool, rawValue)
	f.triggerRecalc(sheet)
	return nil
}

func (f *File) SetCellDate(sheet, cellRef string, value time.Time) error {
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

	_, rawValue := detectCellType(value)
	s.setCellValue(col, row, CellTypeDate, rawValue)
	f.triggerRecalc(sheet)
	return nil
}

func (f *File) GetCellFloat(sheet, cellRef string) (float64, error) {
	if f.closed {
		return 0, ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
		return 0, ErrSheetNotFound
	}

	col, row, err := CellNameToCoordinates(cellRef)
	if err != nil {
		return 0, err
	}

	c := s.getCell(col, row)
	if c == nil {
		return 0, nil
	}

	return parseFloat(c.rawValue)
}

func (f *File) GetCellInt(sheet, cellRef string) (int64, error) {
	if f.closed {
		return 0, ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
		return 0, ErrSheetNotFound
	}

	col, row, err := CellNameToCoordinates(cellRef)
	if err != nil {
		return 0, err
	}

	c := s.getCell(col, row)
	if c == nil {
		return 0, nil
	}

	return parseInt(c.rawValue)
}

func (f *File) GetCellBool(sheet, cellRef string) (bool, error) {
	if f.closed {
		return false, ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
		return false, ErrSheetNotFound
	}

	col, row, err := CellNameToCoordinates(cellRef)
	if err != nil {
		return false, err
	}

	c := s.getCell(col, row)
	if c == nil {
		return false, nil
	}

	return parseBool(c.rawValue)
}

func (f *File) GetCellDate(sheet, cellRef string) (time.Time, error) {
	if f.closed {
		return time.Time{}, ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
		return time.Time{}, ErrSheetNotFound
	}

	col, row, err := CellNameToCoordinates(cellRef)
	if err != nil {
		return time.Time{}, err
	}

	c := s.getCell(col, row)
	if c == nil {
		return time.Time{}, nil
	}

	return parseDate(c.rawValue)
}

func (f *File) GetRows(sheet string) ([][]string, error) {
	if f.closed {
		return nil, ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
		return nil, ErrSheetNotFound
	}

	if s.maxRow == 0 {
		return nil, nil
	}

	result := make([][]string, s.maxRow)
	for rowIdx := 1; rowIdx <= s.maxRow; rowIdx++ {
		rowData := make([]string, s.maxCol)
		r, exists := s.rows[rowIdx]
		if exists {
			for colIdx := 1; colIdx <= s.maxCol; colIdx++ {
				if c, ok := r.cells[colIdx]; ok {
					rowData[colIdx-1] = cellValueToString(c.valueType, c.rawValue)
				}
			}
		}
		result[rowIdx-1] = rowData
	}

	return result, nil
}

func (f *File) SetSheetRow(sheet, startCell string, values []interface{}) error {
	if f.closed {
		return ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
		return ErrSheetNotFound
	}

	col, row, err := CellNameToCoordinates(startCell)
	if err != nil {
		return err
	}

	for i, v := range values {
		valueType, rawValue := detectCellType(v)
		s.setCellValue(col+i, row, valueType, rawValue)
	}

	return nil
}

func (f *File) GetSheetDimension(sheet string) (string, error) {
	if f.closed {
		return "", ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
		return "", ErrSheetNotFound
	}

	if s.maxRow == 0 || s.maxCol == 0 {
		return "A1", nil
	}

	endCell, err := CoordinatesToCellName(s.maxCol, s.maxRow)
	if err != nil {
		return "", err
	}

	return fmt.Sprintf("A1:%s", endCell), nil
}

func (f *File) SetRowValues(sheet string, row int, values []any) error {
	if f.closed {
		return sheetErr(sheet, ErrFileClosed)
	}
	s := f.getSheet(sheet)
	if s == nil {
		return sheetErr(sheet, ErrSheetNotFound)
	}
	if row < 1 {
		return sheetErr(sheet, ErrRowOutOfRange)
	}

	for i, v := range values {
		valueType, rawValue := detectCellType(v)
		s.setCellValue(i+1, row, valueType, rawValue)
	}
	f.triggerRecalc(sheet)
	return nil
}

func (f *File) AppendRows(sheet string, rows [][]any) error {
	if f.closed {
		return sheetErr(sheet, ErrFileClosed)
	}
	s := f.getSheet(sheet)
	if s == nil {
		return sheetErr(sheet, ErrSheetNotFound)
	}

	startRow := s.maxRow + 1
	for i, rowValues := range rows {
		for j, v := range rowValues {
			valueType, rawValue := detectCellType(v)
			s.setCellValue(j+1, startRow+i, valueType, rawValue)
		}
	}
	f.triggerRecalc(sheet)
	return nil
}

func (s *sheet) setCellValue(col, row int, valueType CellType, rawValue string) {
	r := s.getOrCreateRow(row)

	c, ok := r.cells[col]
	if !ok {
		c = &cell{}
		r.cells[col] = c
	}

	c.valueType = valueType
	c.rawValue = rawValue

	if col > s.maxCol {
		s.maxCol = col
	}
	if row > s.maxRow {
		s.maxRow = row
	}
}

func (s *sheet) getCell(col, row int) *cell {
	r, ok := s.rows[row]
	if !ok {
		return nil
	}
	return r.cells[col]
}
