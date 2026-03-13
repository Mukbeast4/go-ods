package goods

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
