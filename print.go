package goods

type printRange struct {
	startCol int
	startRow int
	endCol   int
	endRow   int
}

type PageSetup struct {
	Orientation  string
	PaperWidth   string
	PaperHeight  string
	MarginTop    float64
	MarginBottom float64
	MarginLeft   float64
	MarginRight  float64
}

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
