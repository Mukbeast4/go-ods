package goods

func (f *File) MergeCell(sheet, topLeft, bottomRight string) error {
	if f.closed {
		return ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
		return ErrSheetNotFound
	}

	startCol, startRow, err := CellNameToCoordinates(topLeft)
	if err != nil {
		return err
	}
	endCol, endRow, err := CellNameToCoordinates(bottomRight)
	if err != nil {
		return err
	}

	if startCol > endCol {
		startCol, endCol = endCol, startCol
	}
	if startRow > endRow {
		startRow, endRow = endRow, startRow
	}

	for _, m := range s.merges {
		if rangesOverlap(startCol, startRow, endCol, endRow, m.startCol, m.startRow, m.endCol, m.endRow) {
			return ErrMergeOverlap
		}
	}

	s.merges = append(s.merges, mergeRange{
		startCol: startCol,
		startRow: startRow,
		endCol:   endCol,
		endRow:   endRow,
	})

	r := s.getOrCreateRow(startRow)
	c, ok := r.cells[startCol]
	if !ok {
		c = &cell{valueType: CellTypeEmpty}
		r.cells[startCol] = c
	}
	c.colSpan = endCol - startCol + 1
	c.rowSpan = endRow - startRow + 1

	return nil
}

func (f *File) UnmergeCell(sheet, topLeft, bottomRight string) error {
	if f.closed {
		return ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
		return ErrSheetNotFound
	}

	startCol, startRow, err := CellNameToCoordinates(topLeft)
	if err != nil {
		return err
	}
	endCol, endRow, err := CellNameToCoordinates(bottomRight)
	if err != nil {
		return err
	}

	if startCol > endCol {
		startCol, endCol = endCol, startCol
	}
	if startRow > endRow {
		startRow, endRow = endRow, startRow
	}

	for i, m := range s.merges {
		if m.startCol == startCol && m.startRow == startRow && m.endCol == endCol && m.endRow == endRow {
			s.merges = append(s.merges[:i], s.merges[i+1:]...)

			if r, ok := s.rows[startRow]; ok {
				if c, ok := r.cells[startCol]; ok {
					c.colSpan = 0
					c.rowSpan = 0
				}
			}

			return nil
		}
	}

	return ErrMergeNotFound
}

func (f *File) GetMergeCells(sheet string) ([][2]string, error) {
	if f.closed {
		return nil, ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
		return nil, ErrSheetNotFound
	}

	result := make([][2]string, 0, len(s.merges))
	for _, m := range s.merges {
		start, _ := CoordinatesToCellName(m.startCol, m.startRow)
		end, _ := CoordinatesToCellName(m.endCol, m.endRow)
		result = append(result, [2]string{start, end})
	}

	return result, nil
}

func rangesOverlap(s1c, s1r, e1c, e1r, s2c, s2r, e2c, e2r int) bool {
	return s1c <= e2c && e1c >= s2c && s1r <= e2r && e1r >= s2r
}
