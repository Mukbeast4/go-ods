package goods

type ConditionalFormat struct {
	Range string
	Rules []ConditionalRule
}

type ConditionalRule struct {
	Value           string
	StyleName       string
	BaseCellAddress string
}

func (f *File) SetConditionalFormat(sheet, cellRange string, rules []ConditionalRule) error {
	if f.closed {
		return ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
		return ErrSheetNotFound
	}

	for i, cf := range s.conditionalFormats {
		if cf.Range == cellRange {
			s.conditionalFormats[i].Rules = rules
			return nil
		}
	}

	s.conditionalFormats = append(s.conditionalFormats, ConditionalFormat{
		Range: cellRange,
		Rules: rules,
	})
	return nil
}

func (f *File) GetConditionalFormats(sheet string) ([]ConditionalFormat, error) {
	if f.closed {
		return nil, ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
		return nil, ErrSheetNotFound
	}
	return s.conditionalFormats, nil
}

func (f *File) RemoveConditionalFormat(sheet, cellRange string) error {
	if f.closed {
		return ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
		return ErrSheetNotFound
	}

	for i, cf := range s.conditionalFormats {
		if cf.Range == cellRange {
			s.conditionalFormats = append(s.conditionalFormats[:i], s.conditionalFormats[i+1:]...)
			return nil
		}
	}
	return ErrConditionalFormatNotFound
}
