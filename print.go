package goods

type printRange struct {
	startCol int
	startRow int
	endCol   int
	endRow   int
}

// PageSetup controls the printed layout of a sheet.
//
// Orientation is "portrait" or "landscape". PaperWidth and PaperHeight accept
// ODS dimensions such as "21cm" or "8.5in". Margins are in centimeters.
type PageSetup struct {
	Orientation  string
	PaperWidth   string
	PaperHeight  string
	MarginTop    float64
	MarginBottom float64
	MarginLeft   float64
	MarginRight  float64
}

// SetPrintRange configures the cells printed for the sheet. Any previous range
// is replaced.
func (f *File) SetPrintRange(sheet, topLeft, bottomRight string) error {
	if f.closed {
		return ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
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

	s.printRange = &printRange{
		startCol: sc, startRow: sr,
		endCol: ec, endRow: er,
	}
	return nil
}

// GetPrintRange returns the configured print range as A1-style cells. Both
// results are "" when no range is set.
func (f *File) GetPrintRange(sheet string) (string, string, error) {
	if f.closed {
		return "", "", ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
		return "", "", ErrSheetNotFound
	}

	if s.printRange == nil {
		return "", "", nil
	}

	tl, err := CoordinatesToCellName(s.printRange.startCol, s.printRange.startRow)
	if err != nil {
		return "", "", err
	}
	br, err := CoordinatesToCellName(s.printRange.endCol, s.printRange.endRow)
	if err != nil {
		return "", "", err
	}
	return tl, br, nil
}

// RemovePrintRange clears the print range configured on the sheet.
func (f *File) RemovePrintRange(sheet string) error {
	if f.closed {
		return ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
		return ErrSheetNotFound
	}

	s.printRange = nil
	return nil
}

// SetPageSetup stores a page setup configuration for the sheet. A copy of
// setup is kept internally so later mutations to setup have no effect.
func (f *File) SetPageSetup(sheet string, setup *PageSetup) error {
	if f.closed {
		return ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
		return ErrSheetNotFound
	}

	cp := *setup
	s.pageSetup = &cp
	return nil
}

// GetPageSetup returns a copy of the page setup configured on the sheet, or
// nil when none is set.
func (f *File) GetPageSetup(sheet string) (*PageSetup, error) {
	if f.closed {
		return nil, ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
		return nil, ErrSheetNotFound
	}

	if s.pageSetup == nil {
		return nil, nil
	}

	cp := *s.pageSetup
	return &cp, nil
}
