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

	for i := range f.namedRanges {
		nr := &f.namedRanges[i]
		if nr.sheet == sheet {
			if nr.startCol >= colIdx {
				nr.startCol += count
				nr.endCol += count
			} else if nr.endCol >= colIdx {
				nr.endCol += count
			}
		}
	}

	for i := range f.autoFilters {
		af := &f.autoFilters[i]
		if af.sheet == sheet {
			if af.startCol >= colIdx {
				af.startCol += count
				af.endCol += count
			} else if af.endCol >= colIdx {
				af.endCol += count
			}
		}
	}

	if s.printRange != nil {
		pr := s.printRange
		if pr.startCol >= colIdx {
			pr.startCol += count
			pr.endCol += count
		} else if pr.endCol >= colIdx {
			pr.endCol += count
		}
	}

	f.shiftFormulasOnInsertCols(sheet, colIdx, count)

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

	newNR := make([]namedRange, 0, len(f.namedRanges))
	for _, nr := range f.namedRanges {
		if nr.sheet == sheet {
			if nr.startCol > colIdx {
				nr.startCol--
				nr.endCol--
				newNR = append(newNR, nr)
			} else if nr.endCol < colIdx {
				newNR = append(newNR, nr)
			} else if nr.startCol < colIdx && nr.endCol >= colIdx {
				nr.endCol--
				if nr.endCol >= nr.startCol {
					newNR = append(newNR, nr)
				}
			}
		} else {
			newNR = append(newNR, nr)
		}
	}
	f.namedRanges = newNR

	newAF := make([]autoFilter, 0, len(f.autoFilters))
	for _, af := range f.autoFilters {
		if af.sheet == sheet {
			if af.startCol > colIdx {
				af.startCol--
				af.endCol--
				newAF = append(newAF, af)
			} else if af.endCol < colIdx {
				newAF = append(newAF, af)
			} else if af.startCol < colIdx && af.endCol >= colIdx {
				af.endCol--
				if af.endCol >= af.startCol {
					newAF = append(newAF, af)
				}
			}
		} else {
			newAF = append(newAF, af)
		}
	}
	f.autoFilters = newAF

	if s.printRange != nil {
		pr := s.printRange
		if pr.startCol > colIdx {
			pr.startCol--
			pr.endCol--
		} else if pr.endCol >= colIdx && pr.startCol < colIdx {
			pr.endCol--
			if pr.endCol < pr.startCol {
				s.printRange = nil
			}
		} else if pr.startCol == colIdx && pr.endCol == colIdx {
			s.printRange = nil
		}
	}

	f.shiftFormulasOnRemoveCol(sheet, colIdx)

	return nil
}
