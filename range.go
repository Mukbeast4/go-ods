package goods

type CellRange struct {
	file     *File
	sheet    *sheet
	sName    string
	startCol int
	startRow int
	endCol   int
	endRow   int
	err      error
}

func (f *File) Range(sheet, rangeRef string) *CellRange {
	r := &CellRange{file: f, sName: sheet}

	if f.closed {
		r.err = sheetErr(sheet, ErrFileClosed)
		return r
	}

	s := f.getSheet(sheet)
	if s == nil {
		r.err = sheetErr(sheet, ErrSheetNotFound)
		return r
	}
	r.sheet = s

	startCol, startRow, endCol, endRow, err := splitCellRange(rangeRef)
	if err != nil {
		r.err = sheetErr(sheet, err)
		return r
	}

	if startCol > endCol {
		startCol, endCol = endCol, startCol
	}
	if startRow > endRow {
		startRow, endRow = endRow, startRow
	}

	r.startCol = startCol
	r.startRow = startRow
	r.endCol = endCol
	r.endRow = endRow
	return r
}

func (r *CellRange) SetStyle(styleID int) *CellRange {
	if r.err != nil {
		return r
	}

	if r.file.styles.get(styleID) == nil {
		r.err = sheetErr(r.sName, ErrStyleNotFound)
		return r
	}

	for rowIdx := r.startRow; rowIdx <= r.endRow; rowIdx++ {
		for colIdx := r.startCol; colIdx <= r.endCol; colIdx++ {
			row := r.sheet.getOrCreateRow(rowIdx)
			c, ok := row.cells[colIdx]
			if !ok {
				c = &cell{valueType: CellTypeEmpty}
				row.cells[colIdx] = c
			}
			c.styleID = styleID

			if colIdx > r.sheet.maxCol {
				r.sheet.maxCol = colIdx
			}
		}
		if rowIdx > r.sheet.maxRow {
			r.sheet.maxRow = rowIdx
		}
	}
	return r
}

func (r *CellRange) SetValue(value any) *CellRange {
	if r.err != nil {
		return r
	}

	valueType, rawValue := detectCellType(value)
	for rowIdx := r.startRow; rowIdx <= r.endRow; rowIdx++ {
		for colIdx := r.startCol; colIdx <= r.endCol; colIdx++ {
			r.sheet.setCellValue(colIdx, rowIdx, valueType, rawValue)
		}
	}
	r.file.triggerRecalc(r.sName)
	return r
}

func (r *CellRange) Merge() *CellRange {
	if r.err != nil {
		return r
	}

	for _, m := range r.sheet.merges {
		if rangesOverlap(r.startCol, r.startRow, r.endCol, r.endRow, m.startCol, m.startRow, m.endCol, m.endRow) {
			r.err = sheetErr(r.sName, ErrMergeOverlap)
			return r
		}
	}

	r.sheet.merges = append(r.sheet.merges, mergeRange{
		startCol: r.startCol,
		startRow: r.startRow,
		endCol:   r.endCol,
		endRow:   r.endRow,
	})

	row := r.sheet.getOrCreateRow(r.startRow)
	c, ok := row.cells[r.startCol]
	if !ok {
		c = &cell{valueType: CellTypeEmpty}
		row.cells[r.startCol] = c
	}
	c.colSpan = r.endCol - r.startCol + 1
	c.rowSpan = r.endRow - r.startRow + 1

	return r
}

func (r *CellRange) SetNumberFormat(format string) *CellRange {
	if r.err != nil {
		return r
	}

	for rowIdx := r.startRow; rowIdx <= r.endRow; rowIdx++ {
		for colIdx := r.startCol; colIdx <= r.endCol; colIdx++ {
			row := r.sheet.getOrCreateRow(rowIdx)
			c, ok := row.cells[colIdx]
			if !ok {
				c = &cell{}
				row.cells[colIdx] = c
			}
			c.numberFormat = format

			if colIdx > r.sheet.maxCol {
				r.sheet.maxCol = colIdx
			}
		}
		if rowIdx > r.sheet.maxRow {
			r.sheet.maxRow = rowIdx
		}
	}
	return r
}

func (r *CellRange) Err() error {
	return r.err
}
