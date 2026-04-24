package goods

// Hyperlink binds a URL to a cell. Display is the visible text; if empty the
// URL itself is shown.
type Hyperlink struct {
	URL     string
	Display string
}

// SetCellHyperlink attaches a hyperlink to a cell. When display is empty the
// URL is used. The cell value is set to display if the cell was previously empty.
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

// GetCellHyperlink returns the URL and display text of the hyperlink attached
// to a cell. Both values are "" when no hyperlink is set.
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

// RemoveCellHyperlink clears the hyperlink attached to the cell if any. The
// cell value is left untouched.
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
