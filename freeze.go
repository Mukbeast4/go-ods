package goods

// SetFreezePane pins rows above and columns left of cellRef so they stay
// visible when scrolling. Pass "B2" to freeze the first row and the first column.
func (f *File) SetFreezePane(sheet, cellRef string) error {
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

	s.freezeCol = col - 1
	s.freezeRow = row - 1
	return nil
}

// GetFreezePane returns the frozen column count and row count for the sheet.
// Both are 0 when no pane is frozen.
func (f *File) GetFreezePane(sheet string) (int, int, error) {
	if f.closed {
		return 0, 0, ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
		return 0, 0, ErrSheetNotFound
	}

	return s.freezeCol, s.freezeRow, nil
}

// RemoveFreezePane clears any frozen pane configuration on the sheet.
func (f *File) RemoveFreezePane(sheet string) error {
	if f.closed {
		return ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
		return ErrSheetNotFound
	}

	s.freezeCol = 0
	s.freezeRow = 0
	return nil
}
