package goods

// SetSheetProtection toggles the protected flag on the sheet. Cells whose
// Style.Protected is true become read-only when protection is enabled.
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

// IsSheetProtected reports whether the sheet is currently marked as protected.
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
