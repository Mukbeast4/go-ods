package goods

type Style struct {
	Font      *Font
	Fill      *Fill
	Border    *Border
	Alignment *Alignment
	Protected *bool
}

type Font struct {
	Family        string
	Size          string
	Bold          string
	Italic        string
	Color         string
	Underline     bool
	Strikethrough bool
}

type Fill struct {
	Color string
}

type Border struct {
	Style string
	Width string
	Color string
}

type Alignment struct {
	Horizontal string
	Vertical   string
	WrapText   bool
}

type styleManager struct {
	styles map[int]*Style
	nextID int
}

func newStyleManager() *styleManager {
	return &styleManager{
		styles: make(map[int]*Style),
		nextID: 1,
	}
}

func (sm *styleManager) add(s *Style) int {
	id := sm.nextID
	sm.nextID++
	sm.styles[id] = s
	return id
}

func (sm *styleManager) get(id int) *Style {
	return sm.styles[id]
}

func (f *File) NewStyle(s *Style) (int, error) {
	if f.closed {
		return 0, ErrFileClosed
	}
	if s == nil {
		return 0, ErrStyleNotFound
	}

	id := f.styles.add(s)
	return id, nil
}

func (f *File) SetCellStyle(sheet, topLeft, bottomRight string, styleID int) error {
	if f.closed {
		return ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
		return ErrSheetNotFound
	}

	if f.styles.get(styleID) == nil {
		return ErrStyleNotFound
	}

	startCol, startRow, endCol, endRow, err := splitCellRange(topLeft + ":" + bottomRight)
	if err != nil {
		return err
	}

	for rowIdx := startRow; rowIdx <= endRow; rowIdx++ {
		for colIdx := startCol; colIdx <= endCol; colIdx++ {
			r := s.getOrCreateRow(rowIdx)
			c, ok := r.cells[colIdx]
			if !ok {
				c = &cell{valueType: CellTypeEmpty}
				r.cells[colIdx] = c
			}
			c.styleID = styleID

			if colIdx > s.maxCol {
				s.maxCol = colIdx
			}
		}
		if rowIdx > s.maxRow {
			s.maxRow = rowIdx
		}
	}

	return nil
}
