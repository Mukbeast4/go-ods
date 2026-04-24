package goods

// Comment holds the metadata attached to a cell annotation. Date is the
// ODS dc:date string ("2024-03-15T10:20:30").
type Comment struct {
	Author string
	Date   string
	Text   string
}

// SetCellComment attaches or replaces the comment on a cell.
func (f *File) SetCellComment(sheet, cellRef string, comment *Comment) error {
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
		c = &cell{valueType: CellTypeEmpty}
		r.cells[col] = c
	}

	c.comment = &Comment{
		Author: comment.Author,
		Date:   comment.Date,
		Text:   comment.Text,
	}

	if col > s.maxCol {
		s.maxCol = col
	}
	if row > s.maxRow {
		s.maxRow = row
	}

	return nil
}

// GetCellComment returns a copy of the comment attached to the cell, or nil
// when no comment is set.
func (f *File) GetCellComment(sheet, cellRef string) (*Comment, error) {
	if f.closed {
		return nil, ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
		return nil, ErrSheetNotFound
	}

	col, row, err := CellNameToCoordinates(cellRef)
	if err != nil {
		return nil, err
	}

	c := s.getCell(col, row)
	if c == nil || c.comment == nil {
		return nil, nil
	}

	return &Comment{
		Author: c.comment.Author,
		Date:   c.comment.Date,
		Text:   c.comment.Text,
	}, nil
}

// RemoveCellComment clears the comment attached to the cell if any.
func (f *File) RemoveCellComment(sheet, cellRef string) error {
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
		c.comment = nil
	}

	return nil
}
