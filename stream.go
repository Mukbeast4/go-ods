package goods

import "sort"

type RowIterator struct {
	sheet   *sheet
	rowKeys []int
	current int
}

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

func (it *RowIterator) Next() bool {
	it.current++
	return it.current < len(it.rowKeys)
}

func (it *RowIterator) RowIndex() int {
	if it.current < 0 || it.current >= len(it.rowKeys) {
		return 0
	}
	return it.rowKeys[it.current]
}

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

func (it *RowIterator) Error() error {
	return nil
}
