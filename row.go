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
		if m.startRow >= rowIdx {
			m.startRow += count
			m.endRow += count
		} else if m.endRow >= rowIdx {
			m.endRow += count
		}
	}

	return nil
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

	newMerges := make([]mergeRange, 0, len(s.merges))
	for _, m := range s.merges {
		if m.startRow > rowIdx {
			m.startRow--
			m.endRow--
			newMerges = append(newMerges, m)
		} else if m.endRow < rowIdx {
			newMerges = append(newMerges, m)
		} else if m.startRow < rowIdx && m.endRow >= rowIdx {
			m.endRow--
			if m.endRow >= m.startRow {
				newMerges = append(newMerges, m)
			}
		}
	}
	s.merges = newMerges

	return nil
}
