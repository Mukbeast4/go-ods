package goods

// ConditionalFormat groups a list of ConditionalRule applied to the same cell
// range. Range uses ODS address syntax (e.g. "A1:B10").
type ConditionalFormat struct {
	Range string
	Rules []ConditionalRule
}

// ConditionalRule binds a calcext condition to a style name. Value follows
// calcext syntax such as "value()>10" or "formula-is($A1=TRUE())".
// BaseCellAddress is the anchor cell used to resolve relative references.
type ConditionalRule struct {
	Value           string
	StyleName       string
	BaseCellAddress string
}

// SetConditionalFormat attaches (or replaces) a conditional format on cellRange.
// If a format already exists for the exact range its rules are replaced.
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

// GetConditionalFormats returns every conditional format configured on the sheet.
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

// RemoveConditionalFormat removes the conditional format attached to the exact
// cellRange. It returns ErrConditionalFormatNotFound if no match is found.
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
