package goods

type autoFilter struct {
	sheet    string
	startCol int
	startRow int
	endCol   int
	endRow   int
}

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
