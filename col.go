package goods

func (f *File) SetColWidth(sheet, colName string, width float64) error {
	if f.closed {
		return ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
		return ErrSheetNotFound
	}

	colIdx := columnNameToNumber(colName)
	if colIdx < 1 {
		return ErrColumnOutOfRange
	}

	for len(s.columns) < colIdx {
		s.columns = append(s.columns, column{})
	}

	s.columns[colIdx-1].width = width

	if colIdx > s.maxCol {
		s.maxCol = colIdx
	}

	return nil
}

func (f *File) GetColWidth(sheet, colName string) (float64, error) {
	if f.closed {
		return 0, ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
		return 0, ErrSheetNotFound
	}

	colIdx := columnNameToNumber(colName)
	if colIdx < 1 {
		return 0, ErrColumnOutOfRange
	}

	if colIdx > len(s.columns) {
		return 0, nil
	}

	return s.columns[colIdx-1].width, nil
}

func (f *File) InsertCols(sheet, colName string, count int) error {
	if f.closed {
		return ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
		return ErrSheetNotFound
	}

	colIdx := columnNameToNumber(colName)
	if colIdx < 1 {
		return ErrColumnOutOfRange
	}
	if count < 1 {
		return nil
	}

	newCols := make([]column, 0, len(s.columns)+count)
	for i, c := range s.columns {
		if i+1 == colIdx {
			for range count {
				newCols = append(newCols, column{})
			}
		}
		newCols = append(newCols, c)
	}
	if colIdx > len(s.columns) {
		for range count {
			newCols = append(newCols, column{})
		}
	}
	s.columns = newCols

	for _, r := range s.rows {
		newCells := make(map[int]*cell)
		for idx, c := range r.cells {
			if idx >= colIdx {
				newCells[idx+count] = c
			} else {
				newCells[idx] = c
			}
		}
		r.cells = newCells
	}

	s.maxCol += count

	for i := range s.merges {
		m := &s.merges[i]
		if m.startCol >= colIdx {
			m.startCol += count
			m.endCol += count
		} else if m.endCol >= colIdx {
			m.endCol += count
		}
	}

	return nil
}

func (f *File) RemoveCol(sheet, colName string) error {
	if f.closed {
		return ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
		return ErrSheetNotFound
	}

	colIdx := columnNameToNumber(colName)
	if colIdx < 1 {
		return ErrColumnOutOfRange
	}

	if colIdx <= len(s.columns) {
		s.columns = append(s.columns[:colIdx-1], s.columns[colIdx:]...)
	}

	for _, r := range s.rows {
		delete(r.cells, colIdx)
		newCells := make(map[int]*cell)
		for idx, c := range r.cells {
			if idx > colIdx {
				newCells[idx-1] = c
			} else {
				newCells[idx] = c
			}
		}
		r.cells = newCells
	}

	if s.maxCol > 0 {
		s.maxCol--
	}

	newMerges := make([]mergeRange, 0, len(s.merges))
	for _, m := range s.merges {
		if m.startCol > colIdx {
			m.startCol--
			m.endCol--
			newMerges = append(newMerges, m)
		} else if m.endCol < colIdx {
			newMerges = append(newMerges, m)
		} else if m.startCol < colIdx && m.endCol >= colIdx {
			m.endCol--
			if m.endCol >= m.startCol {
				newMerges = append(newMerges, m)
			}
		}
	}
	s.merges = newMerges

	return nil
}
