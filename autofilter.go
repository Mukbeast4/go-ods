package goods

type autoFilter struct {
	sheet    string
	startCol int
	startRow int
	endCol   int
	endRow   int
	filters  []filterColumn
}

type filterColumn struct {
	colIdx int
	values []string
}

// FilterCriteria describes which values are kept visible for a given column of
// an auto-filter. Column is 1-based. Rows whose value in that column is not in
// Values are hidden.
type FilterCriteria struct {
	Column int
	Values []string
}

// SortKey describes one key used by SetSort. Column is 1-based.
type SortKey struct {
	Column     int
	Descending bool
}

// SetAutoFilter enables an auto-filter over the given range. Only one auto-filter
// may exist per sheet; a second call replaces the previous range.
func (f *File) SetAutoFilter(sheet, topLeft, bottomRight string) error {
	if f.closed {
		return ErrFileClosed
	}
	if f.getSheet(sheet) == nil {
		return ErrSheetNotFound
	}

	sc, sr, err := CellNameToCoordinates(topLeft)
	if err != nil {
		return err
	}
	ec, er, err := CellNameToCoordinates(bottomRight)
	if err != nil {
		return err
	}

	for i, af := range f.autoFilters {
		if af.sheet == sheet {
			f.autoFilters[i] = autoFilter{
				sheet: sheet, startCol: sc, startRow: sr,
				endCol: ec, endRow: er,
			}
			return nil
		}
	}

	f.autoFilters = append(f.autoFilters, autoFilter{
		sheet: sheet, startCol: sc, startRow: sr,
		endCol: ec, endRow: er,
	})
	return nil
}

// GetAutoFilter returns the top-left and bottom-right A1-style cells of the
// sheet's auto-filter. It returns ErrAutoFilterNotFound when none is set.
func (f *File) GetAutoFilter(sheet string) (string, string, error) {
	if f.closed {
		return "", "", ErrFileClosed
	}
	if f.getSheet(sheet) == nil {
		return "", "", ErrSheetNotFound
	}

	for _, af := range f.autoFilters {
		if af.sheet == sheet {
			tl, err := CoordinatesToCellName(af.startCol, af.startRow)
			if err != nil {
				return "", "", err
			}
			br, err := CoordinatesToCellName(af.endCol, af.endRow)
			if err != nil {
				return "", "", err
			}
			return tl, br, nil
		}
	}

	return "", "", ErrAutoFilterNotFound
}

// RemoveAutoFilter clears the sheet's auto-filter including any criteria.
func (f *File) RemoveAutoFilter(sheet string) error {
	if f.closed {
		return ErrFileClosed
	}
	if f.getSheet(sheet) == nil {
		return ErrSheetNotFound
	}

	for i, af := range f.autoFilters {
		if af.sheet == sheet {
			f.autoFilters = append(f.autoFilters[:i], f.autoFilters[i+1:]...)
			return nil
		}
	}
	return ErrAutoFilterNotFound
}

// SetFilterCriteria replaces the filtering criteria attached to the sheet's
// auto-filter. An auto-filter must already be set on the sheet.
func (f *File) SetFilterCriteria(sheet string, criteria []FilterCriteria) error {
	if f.closed {
		return ErrFileClosed
	}
	if f.getSheet(sheet) == nil {
		return ErrSheetNotFound
	}

	for i, af := range f.autoFilters {
		if af.sheet == sheet {
			f.autoFilters[i].filters = nil
			for _, c := range criteria {
				f.autoFilters[i].filters = append(f.autoFilters[i].filters, filterColumn{
					colIdx: c.Column,
					values: c.Values,
				})
			}
			return nil
		}
	}
	return ErrAutoFilterNotFound
}

// GetFilterCriteria returns the criteria attached to the sheet's auto-filter.
func (f *File) GetFilterCriteria(sheet string) ([]FilterCriteria, error) {
	if f.closed {
		return nil, ErrFileClosed
	}
	if f.getSheet(sheet) == nil {
		return nil, ErrSheetNotFound
	}

	for _, af := range f.autoFilters {
		if af.sheet == sheet {
			var result []FilterCriteria
			for _, fc := range af.filters {
				result = append(result, FilterCriteria{
					Column: fc.colIdx,
					Values: fc.values,
				})
			}
			return result, nil
		}
	}
	return nil, ErrAutoFilterNotFound
}

// ClearFilterCriteria removes every criterion while keeping the auto-filter range.
func (f *File) ClearFilterCriteria(sheet string) error {
	if f.closed {
		return ErrFileClosed
	}
	if f.getSheet(sheet) == nil {
		return ErrSheetNotFound
	}

	for i, af := range f.autoFilters {
		if af.sheet == sheet {
			f.autoFilters[i].filters = nil
			return nil
		}
	}
	return ErrAutoFilterNotFound
}

// SetSort stores a primary/secondary sort configuration for the sheet. The
// keys are applied in order.
func (f *File) SetSort(sheet string, keys []SortKey) error {
	if f.closed {
		return ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
		return ErrSheetNotFound
	}
	s.sortKeys = keys
	return nil
}

// GetSort returns the sort configuration stored for the sheet.
func (f *File) GetSort(sheet string) ([]SortKey, error) {
	if f.closed {
		return nil, ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
		return nil, ErrSheetNotFound
	}
	return s.sortKeys, nil
}

// RemoveSort clears the sort configuration stored for the sheet.
func (f *File) RemoveSort(sheet string) error {
	if f.closed {
		return ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
		return ErrSheetNotFound
	}
	s.sortKeys = nil
	return nil
}
