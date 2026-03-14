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
		s.columns = append(s.columns, column{visible: true})
	}

	s.columns[colIdx-1].width = width

	if colIdx > s.maxCol {
		s.maxCol = colIdx
	}

	return nil
}

func (f *File) SetColVisible(sheet, colName string, visible bool) error {
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
		s.columns = append(s.columns, column{visible: true})
	}

	s.columns[colIdx-1].visible = visible
	return nil
}

func (f *File) GetColVisible(sheet, colName string) (bool, error) {
	if f.closed {
		return false, ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
		return false, ErrSheetNotFound
	}

	colIdx := columnNameToNumber(colName)
	if colIdx < 1 {
		return false, ErrColumnOutOfRange
	}

	if colIdx > len(s.columns) {
		return true, nil
	}

	return s.columns[colIdx-1].visible, nil
}

func (f *File) SetColAutoFit(sheet, colName string, autoFit bool) error {
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
		s.columns = append(s.columns, column{visible: true})
	}

	s.columns[colIdx-1].autoFit = autoFit
	return nil
}

func (f *File) GetColAutoFit(sheet, colName string) (bool, error) {
	if f.closed {
		return false, ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
		return false, ErrSheetNotFound
	}

	colIdx := columnNameToNumber(colName)
	if colIdx < 1 {
		return false, ErrColumnOutOfRange
	}

	if colIdx > len(s.columns) {
		return false, nil
	}

	return s.columns[colIdx-1].autoFit, nil
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

func shiftColRange(startCol *int, endCol *int, insertAt int, count int) {
	if *startCol >= insertAt {
		*startCol += count
		*endCol += count
	} else if *endCol >= insertAt {
		*endCol += count
	}
}

func shrinkColRangeOnRemove(startCol, endCol, removeAt int) (int, int, bool) {
	if startCol > removeAt {
		return startCol - 1, endCol - 1, true
	}
	if endCol < removeAt {
		return startCol, endCol, true
	}
	if startCol < removeAt && endCol >= removeAt {
		endCol--
		return startCol, endCol, endCol >= startCol
	}
	return startCol, endCol, false
}

func insertBlankColumns(cols []column, colIdx, count int) []column {
	newCols := make([]column, 0, len(cols)+count)
	for i, c := range cols {
		if i+1 == colIdx {
			for range count {
				newCols = append(newCols, column{visible: true})
			}
		}
		newCols = append(newCols, c)
	}
	if colIdx > len(cols) {
		for range count {
			newCols = append(newCols, column{visible: true})
		}
	}
	return newCols
}

func shiftCellsOnInsertCol(s *sheet, colIdx, count int) {
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

	s.columns = insertBlankColumns(s.columns, colIdx, count)
	shiftCellsOnInsertCol(s, colIdx, count)
	s.maxCol += count

	for i := range s.merges {
		m := &s.merges[i]
		shiftColRange(&m.startCol, &m.endCol, colIdx, count)
	}

	for i := range f.namedRanges {
		nr := &f.namedRanges[i]
		if nr.sheet == sheet {
			shiftColRange(&nr.startCol, &nr.endCol, colIdx, count)
		}
	}

	for i := range f.autoFilters {
		af := &f.autoFilters[i]
		if af.sheet == sheet {
			shiftColRange(&af.startCol, &af.endCol, colIdx, count)
		}
	}

	if s.printRange != nil {
		shiftColRange(&s.printRange.startCol, &s.printRange.endCol, colIdx, count)
	}

	f.shiftFormulasOnInsertCols(sheet, colIdx, count)

	return nil
}

func removeCellsAtCol(s *sheet, colIdx int) {
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
}

func removeColumnFromSlice(cols []column, colIdx int) []column {
	if colIdx <= len(cols) {
		return append(cols[:colIdx-1], cols[colIdx:]...)
	}
	return cols
}

func shrinkMergesOnRemoveCol(merges []mergeRange, colIdx int) []mergeRange {
	result := make([]mergeRange, 0, len(merges))
	for _, m := range merges {
		newStart, newEnd, keep := shrinkColRangeOnRemove(m.startCol, m.endCol, colIdx)
		if keep {
			m.startCol = newStart
			m.endCol = newEnd
			result = append(result, m)
		}
	}
	return result
}

func shrinkNamedRangesOnRemoveCol(ranges []namedRange, sheet string, colIdx int) []namedRange {
	result := make([]namedRange, 0, len(ranges))
	for _, nr := range ranges {
		if nr.sheet != sheet {
			result = append(result, nr)
			continue
		}
		newStart, newEnd, keep := shrinkColRangeOnRemove(nr.startCol, nr.endCol, colIdx)
		if keep {
			nr.startCol = newStart
			nr.endCol = newEnd
			result = append(result, nr)
		}
	}
	return result
}

func shrinkAutoFiltersOnRemoveCol(filters []autoFilter, sheet string, colIdx int) []autoFilter {
	result := make([]autoFilter, 0, len(filters))
	for _, af := range filters {
		if af.sheet != sheet {
			result = append(result, af)
			continue
		}
		newStart, newEnd, keep := shrinkColRangeOnRemove(af.startCol, af.endCol, colIdx)
		if keep {
			af.startCol = newStart
			af.endCol = newEnd
			result = append(result, af)
		}
	}
	return result
}

func shrinkPrintRangeOnRemoveCol(s *sheet, colIdx int) {
	if s.printRange == nil {
		return
	}
	pr := s.printRange
	newStart, newEnd, keep := shrinkColRangeOnRemove(pr.startCol, pr.endCol, colIdx)
	if keep {
		pr.startCol = newStart
		pr.endCol = newEnd
	} else {
		s.printRange = nil
	}
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

	s.columns = removeColumnFromSlice(s.columns, colIdx)
	removeCellsAtCol(s, colIdx)

	if s.maxCol > 0 {
		s.maxCol--
	}

	s.merges = shrinkMergesOnRemoveCol(s.merges, colIdx)
	f.namedRanges = shrinkNamedRangesOnRemoveCol(f.namedRanges, sheet, colIdx)
	f.autoFilters = shrinkAutoFiltersOnRemoveCol(f.autoFilters, sheet, colIdx)
	shrinkPrintRangeOnRemoveCol(s, colIdx)

	f.shiftFormulasOnRemoveCol(sheet, colIdx)

	return nil
}
