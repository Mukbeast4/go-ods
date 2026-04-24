package goods

// Style describes the visual formatting applied to one or more cells. Any
// nested field left nil is ignored when the style is resolved. Set Protected
// to &true to participate in sheet protection locking.
type Style struct {
	Font      *Font
	Fill      *Fill
	Border    *Border
	Alignment *Alignment
	Protected *bool
}

// Font configures text rendering for a cell. Size follows ODS conventions such
// as "12pt". VerticalAlign accepts "super" or "sub" for super/subscript. Bold
// and Italic use string form to allow three states ("", "bold", "normal").
type Font struct {
	Family             string
	Size               string
	Bold               string
	Italic             string
	Color              string
	Underline          bool
	Strikethrough      bool
	StrikethroughColor string
	VerticalAlign      string
}

// Fill describes the background color of a cell using a hex value such as "#FFEE00".
type Fill struct {
	Color string
}

// Border defines a single border line applied uniformly to all four sides.
// Style follows ODS syntax ("solid", "dashed"...), Width uses units like "0.5pt".
type Border struct {
	Style string
	Width string
	Color string
}

// Alignment controls how content is laid out inside a cell. Rotation is in
// degrees (0-360) and Indent is expressed in 0.25cm increments.
type Alignment struct {
	Horizontal string
	Vertical   string
	WrapText   bool
	Rotation   int
	Indent     int
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

// NewStyle registers a style in the workbook's style manager and returns an
// opaque id to pass to SetCellStyle. It returns ErrStyleNotFound when s is nil.
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

// SetCellStyle applies the style identified by styleID to every cell in the
// rectangular range between topLeft and bottomRight (inclusive). Empty cells in
// the range are created so the style persists after save.
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
