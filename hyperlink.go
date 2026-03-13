package goods

type Hyperlink struct {
	URL     string
	Display string
}

func (f *File) SetCellHyperlink(sheet, cellRef, url, display string) error {
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
		c = &cell{valueType: CellTypeString}
		r.cells[col] = c
	}

	if display == "" {
		display = url
	}

	c.hyperlink = &Hyperlink{URL: url, Display: display}
	if c.rawValue == "" {
		c.rawValue = display
		c.valueType = CellTypeString
	}

	if col > s.maxCol {
		s.maxCol = col
	}
	if row > s.maxRow {
		s.maxRow = row
	}

	return nil
}

func (f *File) GetCellHyperlink(sheet, cellRef string) (string, string, error) {
	if f.closed {
		return "", "", ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
		return "", "", ErrSheetNotFound
	}

	col, row, err := CellNameToCoordinates(cellRef)
	if err != nil {
		return "", "", err
	}

	c := s.getCell(col, row)
	if c == nil || c.hyperlink == nil {
		return "", "", nil
	}

	return c.hyperlink.URL, c.hyperlink.Display, nil
}

func (f *File) RemoveCellHyperlink(sheet, cellRef string) error {
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

	c := s.getCell(col, row)
	if c != nil {
		c.hyperlink = nil
	}

	return nil
}
