package goods

func (f *File) SetRowHeight(sheet string, rowIdx int, height float64) error {
	if f.closed {
		return ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
		return ErrSheetNotFound
	}
	if rowIdx < 1 {
		return ErrRowOutOfRange
	}

	r := s.getOrCreateRow(rowIdx)
	r.height = height

	if rowIdx > s.maxRow {
		s.maxRow = rowIdx
	}

	return nil
}

func (f *File) SetRowVisible(sheet string, rowIdx int, visible bool) error {
	if f.closed {
		return ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
		return ErrSheetNotFound
	}
	if rowIdx < 1 {
		return ErrRowOutOfRange
	}

	r := s.getOrCreateRow(rowIdx)
	r.visible = visible

	if rowIdx > s.maxRow {
		s.maxRow = rowIdx
	}

	return nil
}

func (f *File) GetRowVisible(sheet string, rowIdx int) (bool, error) {
	if f.closed {
		return false, ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
		return false, ErrSheetNotFound
	}
	if rowIdx < 1 {
		return false, ErrRowOutOfRange
	}

	r, ok := s.rows[rowIdx]
	if !ok {
		return true, nil
	}

	return r.visible, nil
}

func (f *File) GetRowHeight(sheet string, rowIdx int) (float64, error) {
	if f.closed {
		return 0, ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
		return 0, ErrSheetNotFound
	}
	if rowIdx < 1 {
		return 0, ErrRowOutOfRange
	}

	r, ok := s.rows[rowIdx]
	if !ok {
		return 0, nil
	}

	return r.height, nil
}

func shiftRowRange(startRow *int, endRow *int, insertAt int, count int) {
	if *startRow >= insertAt {
		*startRow += count
		*endRow += count
	} else if *endRow >= insertAt {
		*endRow += count
	}
}

func shrinkRowRangeOnRemove(startRow, endRow, removeAt int) (int, int, bool) {
	if startRow > removeAt {
		return startRow - 1, endRow - 1, true
	}
	if endRow < removeAt {
		return startRow, endRow, true
	}
	if startRow < removeAt && endRow >= removeAt {
		endRow--
		return startRow, endRow, endRow >= startRow
	}
	return startRow, endRow, false
}

func (f *File) InsertRows(sheet string, rowIdx, count int) error {
	if f.closed {
		return ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
		return ErrSheetNotFound
	}
	if rowIdx < 1 {
		return ErrRowOutOfRange
	}
	if count < 1 {
		return nil
	}

	newRows := make(map[int]*row)
	for idx, r := range s.rows {
		if idx >= rowIdx {
			newRows[idx+count] = r
		} else {
			newRows[idx] = r
		}
	}
	s.rows = newRows
	s.maxRow += count

	for i := range s.merges {
		m := &s.merges[i]
		shiftRowRange(&m.startRow, &m.endRow, rowIdx, count)
	}

	for i := range f.namedRanges {
		nr := &f.namedRanges[i]
		if nr.sheet == sheet {
			shiftRowRange(&nr.startRow, &nr.endRow, rowIdx, count)
		}
	}

	for i := range f.autoFilters {
		af := &f.autoFilters[i]
		if af.sheet == sheet {
			shiftRowRange(&af.startRow, &af.endRow, rowIdx, count)
		}
	}

	if s.printRange != nil {
		shiftRowRange(&s.printRange.startRow, &s.printRange.endRow, rowIdx, count)
	}

	f.shiftFormulasOnInsertRows(sheet, rowIdx, count)

	return nil
}

func removeRowAndShift(s *sheet, rowIdx int) {
	delete(s.rows, rowIdx)

	newRows := make(map[int]*row)
	for idx, r := range s.rows {
		if idx > rowIdx {
			newRows[idx-1] = r
		} else {
			newRows[idx] = r
		}
	}
	s.rows = newRows

	if s.maxRow > 0 {
		s.maxRow--
	}
}

func shrinkMergesOnRemoveRow(s *sheet, rowIdx int) {
	newMerges := make([]mergeRange, 0, len(s.merges))
	for _, m := range s.merges {
		newStart, newEnd, keep := shrinkRowRangeOnRemove(m.startRow, m.endRow, rowIdx)
		if keep {
			m.startRow = newStart
			m.endRow = newEnd
			newMerges = append(newMerges, m)
		}
	}
	s.merges = newMerges
}

func shrinkNamedRangesOnRemoveRow(f *File, sheet string, rowIdx int) {
	newNR := make([]namedRange, 0, len(f.namedRanges))
	for _, nr := range f.namedRanges {
		if nr.sheet != sheet {
			newNR = append(newNR, nr)
			continue
		}
		newStart, newEnd, keep := shrinkRowRangeOnRemove(nr.startRow, nr.endRow, rowIdx)
		if keep {
			nr.startRow = newStart
			nr.endRow = newEnd
			newNR = append(newNR, nr)
		}
	}
	f.namedRanges = newNR
}

func shrinkAutoFiltersOnRemoveRow(f *File, sheet string, rowIdx int) {
	newAF := make([]autoFilter, 0, len(f.autoFilters))
	for _, af := range f.autoFilters {
		if af.sheet != sheet {
			newAF = append(newAF, af)
			continue
		}
		newStart, newEnd, keep := shrinkRowRangeOnRemove(af.startRow, af.endRow, rowIdx)
		if keep {
			af.startRow = newStart
			af.endRow = newEnd
			newAF = append(newAF, af)
		}
	}
	f.autoFilters = newAF
}

func shrinkPrintRangeOnRemoveRow(s *sheet, rowIdx int) {
	if s.printRange == nil {
		return
	}
	pr := s.printRange
	newStart, newEnd, keep := shrinkRowRangeOnRemove(pr.startRow, pr.endRow, rowIdx)
	if keep {
		pr.startRow = newStart
		pr.endRow = newEnd
	} else {
		s.printRange = nil
	}
}

func (f *File) RemoveRow(sheet string, rowIdx int) error {
	if f.closed {
		return ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
		return ErrSheetNotFound
	}
	if rowIdx < 1 {
		return ErrRowOutOfRange
	}

	removeRowAndShift(s, rowIdx)
	shrinkMergesOnRemoveRow(s, rowIdx)
	shrinkNamedRangesOnRemoveRow(f, sheet, rowIdx)
	shrinkAutoFiltersOnRemoveRow(f, sheet, rowIdx)
	shrinkPrintRangeOnRemoveRow(s, rowIdx)
	f.shiftFormulasOnRemoveRow(sheet, rowIdx)

	return nil
}
