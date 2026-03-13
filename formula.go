package goods

func (f *File) SetCellFormula(sheet, cellRef, formula string) error {
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
		c = &cell{valueType: CellTypeFloat}
		r.cells[col] = c
	}
	c.formula = formula

	if col > s.maxCol {
		s.maxCol = col
	}
	if row > s.maxRow {
		s.maxRow = row
	}

	return nil
}

func (f *File) GetCellFormula(sheet, cellRef string) (string, error) {
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

	return c.formula, nil
}
