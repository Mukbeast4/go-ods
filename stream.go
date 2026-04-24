package goods

import "sort"

// RowIterator yields rows from a sheet in ascending row-index order. Empty
// rows are skipped. Use it to stream large sheets without materializing the
// entire document with GetRows.
type RowIterator struct {
	sheet   *sheet
	rowKeys []int
	current int
}

// NewRowIterator returns an iterator over the non-empty rows of the sheet.
func (f *File) NewRowIterator(sheet string) (*RowIterator, error) {
	if f.closed {
		return nil, ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
		return nil, ErrSheetNotFound
	}

	keys := make([]int, 0, len(s.rows))
	for k := range s.rows {
		keys = append(keys, k)
	}
	sort.Ints(keys)

	return &RowIterator{
		sheet:   s,
		rowKeys: keys,
		current: -1,
	}, nil
}

// Next advances to the next non-empty row. It returns false when the iterator
// is exhausted.
func (it *RowIterator) Next() bool {
	it.current++
	return it.current < len(it.rowKeys)
}

// RowIndex returns the 1-based row index of the row currently yielded by the
// iterator. It returns 0 before Next has been called or after it returned false.
func (it *RowIterator) RowIndex() int {
	if it.current < 0 || it.current >= len(it.rowKeys) {
		return 0
	}
	return it.rowKeys[it.current]
}

// Row returns the current row as a dense slice of strings sized by the sheet's
// max column. Empty cells yield "".
func (it *RowIterator) Row() []string {
	if it.current < 0 || it.current >= len(it.rowKeys) {
		return nil
	}

	rowIdx := it.rowKeys[it.current]
	r, ok := it.sheet.rows[rowIdx]
	if !ok {
		return nil
	}

	maxCol := it.sheet.maxCol
	result := make([]string, maxCol)
	for colIdx := 1; colIdx <= maxCol; colIdx++ {
		if c, ok := r.cells[colIdx]; ok {
			result[colIdx-1] = cellValueToString(c.valueType, c.rawValue)
		}
	}

	return result
}

// Error reports the first error encountered during iteration. Currently
// iteration is infallible and Error always returns nil.
func (it *RowIterator) Error() error {
	return nil
}
