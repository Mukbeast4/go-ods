package goods

func (f *File) SetSheetProtection(sheet string, protected bool) error {
	if f.closed {
		return ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
		return ErrSheetNotFound
	}
	s.protected = protected
	return nil
}

func (f *File) IsSheetProtected(sheet string) (bool, error) {
	if f.closed {
		return false, ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
		return false, ErrSheetNotFound
	}
	return s.protected, nil
}
