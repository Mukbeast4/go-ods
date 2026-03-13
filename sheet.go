package goods

func (f *File) NewSheet(name string) (int, error) {
	if f.closed {
		return -1, ErrFileClosed
	}
	if name == "" {
		return -1, ErrSheetNameEmpty
	}
	if f.getSheet(name) != nil {
		return -1, ErrSheetExists
	}

	s := &sheet{
		name:    name,
		rows:    make(map[int]*row),
		columns: make([]column, 0),
	}
	f.sheets = append(f.sheets, s)
	return len(f.sheets) - 1, nil
}

func (f *File) DeleteSheet(name string) error {
	if f.closed {
		return ErrFileClosed
	}
	if len(f.sheets) <= 1 {
		return ErrNoSheets
	}

	for i, s := range f.sheets {
		if s.name == name {
			f.sheets = append(f.sheets[:i], f.sheets[i+1:]...)
			if f.activeSheet >= len(f.sheets) {
				f.activeSheet = len(f.sheets) - 1
			}

			newNR := make([]namedRange, 0, len(f.namedRanges))
			for _, nr := range f.namedRanges {
				if nr.sheet != name {
					newNR = append(newNR, nr)
				}
			}
			f.namedRanges = newNR

			newAF := make([]autoFilter, 0, len(f.autoFilters))
			for _, af := range f.autoFilters {
				if af.sheet != name {
					newAF = append(newAF, af)
				}
			}
			f.autoFilters = newAF

			return nil
		}
	}
	return ErrSheetNotFound
}

func (f *File) GetSheetList() []string {
	names := make([]string, len(f.sheets))
	for i, s := range f.sheets {
		names[i] = s.name
	}
	return names
}

func (f *File) GetSheetName(index int) (string, error) {
	if index < 0 || index >= len(f.sheets) {
		return "", ErrSheetNotFound
	}
	return f.sheets[index].name, nil
}

func (f *File) SetSheetName(oldName, newName string) error {
	if f.closed {
		return ErrFileClosed
	}
	if newName == "" {
		return ErrSheetNameEmpty
	}
	if f.getSheet(newName) != nil {
		return ErrSheetExists
	}

	s := f.getSheet(oldName)
	if s == nil {
		return ErrSheetNotFound
	}

	s.name = newName

	for i := range f.namedRanges {
		if f.namedRanges[i].sheet == oldName {
			f.namedRanges[i].sheet = newName
		}
	}
	for i := range f.autoFilters {
		if f.autoFilters[i].sheet == oldName {
			f.autoFilters[i].sheet = newName
		}
	}

	return nil
}

func (f *File) GetActiveSheetIndex() int {
	return f.activeSheet
}

func (f *File) SetActiveSheet(index int) error {
	if index < 0 || index >= len(f.sheets) {
		return ErrSheetNotFound
	}
	f.activeSheet = index
	return nil
}

func (f *File) SheetCount() int {
	return len(f.sheets)
}

func (f *File) GetSheetIndex(name string) (int, error) {
	for i, s := range f.sheets {
		if s.name == name {
			return i, nil
		}
	}
	return -1, ErrSheetNotFound
}

func (f *File) CopySheet(source, target string) error {
	if f.closed {
		return ErrFileClosed
	}
	if target == "" {
		return ErrSheetNameEmpty
	}
	if f.getSheet(target) != nil {
		return ErrSheetExists
	}

	src := f.getSheet(source)
	if src == nil {
		return ErrSheetNotFound
	}

	dst := &sheet{
		name:      target,
		columns:   make([]column, len(src.columns)),
		rows:      make(map[int]*row),
		maxRow:    src.maxRow,
		maxCol:    src.maxCol,
		freezeCol: src.freezeCol,
		freezeRow: src.freezeRow,
	}

	copy(dst.columns, src.columns)

	for idx, r := range src.rows {
		newRow := &row{
			cells:   make(map[int]*cell),
			height:  r.height,
			visible: r.visible,
		}
		for cIdx, c := range r.cells {
			newCell := *c
			if c.comment != nil {
				cp := *c.comment
				newCell.comment = &cp
			}
			if c.hyperlink != nil {
				cp := *c.hyperlink
				newCell.hyperlink = &cp
			}
			newRow.cells[cIdx] = &newCell
		}
		dst.rows[idx] = newRow
	}

	dst.merges = make([]mergeRange, len(src.merges))
	copy(dst.merges, src.merges)

	dst.validations = make([]*dataValidation, len(src.validations))
	for i, v := range src.validations {
		dvCopy := *v.validation
		dst.validations[i] = &dataValidation{
			name:       v.name,
			validation: &dvCopy,
		}
	}

	if src.printRange != nil {
		cp := *src.printRange
		dst.printRange = &cp
	}
	if src.pageSetup != nil {
		cp := *src.pageSetup
		dst.pageSetup = &cp
	}

	f.sheets = append(f.sheets, dst)
	return nil
}
