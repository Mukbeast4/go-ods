package goods

import (
	"fmt"
	"strings"
)

type namedRange struct {
	name     string
	sheet    string
	startCol int
	startRow int
	endCol   int
	endRow   int
}

type NamedRangeInfo struct {
	Name        string
	Sheet       string
	TopLeft     string
	BottomRight string
}

func (f *File) SetNamedRange(name, sheet, topLeft, bottomRight string) error {
	if f.closed {
		return ErrFileClosed
	}
	if name == "" {
		return ErrInvalidCell
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

	for i, nr := range f.namedRanges {
		if nr.name == name {
			f.namedRanges[i] = namedRange{
				name: name, sheet: sheet,
				startCol: sc, startRow: sr,
				endCol: ec, endRow: er,
			}
			return nil
		}
	}

	f.namedRanges = append(f.namedRanges, namedRange{
		name: name, sheet: sheet,
		startCol: sc, startRow: sr,
		endCol: ec, endRow: er,
	})
	return nil
}

func (f *File) GetNamedRange(name string) (*NamedRangeInfo, error) {
	if f.closed {
		return nil, ErrFileClosed
	}

	for _, nr := range f.namedRanges {
		if nr.name == name {
			return namedRangeToInfo(&nr)
		}
	}
	return nil, ErrNamedRangeNotFound
}

func (f *File) DeleteNamedRange(name string) error {
	if f.closed {
		return ErrFileClosed
	}

	for i, nr := range f.namedRanges {
		if nr.name == name {
			f.namedRanges = append(f.namedRanges[:i], f.namedRanges[i+1:]...)
			return nil
		}
	}
	return ErrNamedRangeNotFound
}

func (f *File) GetNamedRanges() []NamedRangeInfo {
	result := make([]NamedRangeInfo, 0, len(f.namedRanges))
	for _, nr := range f.namedRanges {
		info, err := namedRangeToInfo(&nr)
		if err == nil {
			result = append(result, *info)
		}
	}
	return result
}

func namedRangeToInfo(nr *namedRange) (*NamedRangeInfo, error) {
	tl, err := CoordinatesToCellName(nr.startCol, nr.startRow)
	if err != nil {
		return nil, err
	}
	br, err := CoordinatesToCellName(nr.endCol, nr.endRow)
	if err != nil {
		return nil, err
	}
	return &NamedRangeInfo{
		Name:        nr.name,
		Sheet:       nr.sheet,
		TopLeft:     tl,
		BottomRight: br,
	}, nil
}

func formatODSCellAddress(sheet string, col, row int) string {
	colName := columnNumberToName(col)
	return fmt.Sprintf("$%s.$%s$%d", sheet, colName, row)
}

func formatODSRangeAddress(sheet string, sc, sr, ec, er int) string {
	return formatODSCellAddress(sheet, sc, sr) + ":" + formatODSCellAddress(sheet, ec, er)
}

func parseODSCellAddress(addr string) (sheet string, col, row int, err error) {
	addr = strings.ReplaceAll(addr, "$", "")

	dotIdx := strings.LastIndex(addr, ".")
	if dotIdx < 0 {
		err = ErrInvalidCell
		return
	}

	sheet = addr[:dotIdx]
	cellRef := addr[dotIdx+1:]

	col, row, err = CellNameToCoordinates(cellRef)
	return
}

func parseODSRangeAddress(addr string) (sheet string, sc, sr, ec, er int, err error) {
	parts := strings.SplitN(addr, ":", 2)
	if len(parts) != 2 {
		sheet, sc, sr, err = parseODSCellAddress(addr)
		ec, er = sc, sr
		return
	}

	sheet, sc, sr, err = parseODSCellAddress(parts[0])
	if err != nil {
		return
	}

	_, ec, er, err = parseODSCellAddress(parts[1])
	return
}
