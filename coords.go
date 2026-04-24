package goods

import (
	"fmt"
	"strings"
)

func parseRowNumber(s string) (int, error) {
	if s == "" {
		return 0, ErrInvalidCell
	}
	row := 0
	for _, c := range s {
		if c < '0' || c > '9' {
			return 0, ErrInvalidCell
		}
		row = row*10 + int(c-'0')
	}
	if row < 1 {
		return 0, ErrInvalidCell
	}
	return row, nil
}

// CellNameToCoordinates converts an A1-style cell reference such as "B12" into
// 1-based column and row numbers. It returns ErrInvalidCell when the reference
// is malformed.
func CellNameToCoordinates(cell string) (col, row int, err error) {
	cell = strings.TrimSpace(cell)
	if cell == "" {
		return 0, 0, ErrInvalidCell
	}

	colStr := ""
	rowStr := ""
	for i, c := range cell {
		if c >= 'A' && c <= 'Z' || c >= 'a' && c <= 'z' {
			colStr += string(c)
		} else if c >= '0' && c <= '9' {
			rowStr = cell[i:]
			break
		} else {
			return 0, 0, ErrInvalidCell
		}
	}

	if colStr == "" || rowStr == "" {
		return 0, 0, ErrInvalidCell
	}

	col = columnNameToNumber(strings.ToUpper(colStr))
	if col < 1 {
		return 0, 0, ErrInvalidCell
	}

	row, err = parseRowNumber(rowStr)
	if err != nil {
		return 0, 0, err
	}

	return col, row, nil
}

// CoordinatesToCellName converts 1-based column and row numbers into an
// A1-style cell reference. It returns ErrInvalidCoords when either argument is
// below 1.
func CoordinatesToCellName(col, row int) (string, error) {
	if col < 1 || row < 1 {
		return "", ErrInvalidCoords
	}
	return fmt.Sprintf("%s%d", columnNumberToName(col), row), nil
}

func columnNameToNumber(name string) int {
	result := 0
	for _, c := range name {
		result = result*26 + int(c-'A') + 1
	}
	return result
}

func columnNumberToName(col int) string {
	result := ""
	for col > 0 {
		col--
		result = string(rune('A'+col%26)) + result
		col /= 26
	}
	return result
}

// Cell builds an A1-style cell reference from 1-based coordinates. It panics
// when col or row is below 1; use CoordinatesToCellName for a non-panicking
// alternative.
func Cell(col, row int) string {
	name, err := CoordinatesToCellName(col, row)
	if err != nil {
		panic(fmt.Sprintf("goods.Cell(%d, %d): %v", col, row, err))
	}
	return name
}

// Cells builds an A1-style range reference such as "A1:C10" from two pairs of
// 1-based coordinates.
func Cells(startCol, startRow, endCol, endRow int) string {
	start := Cell(startCol, startRow)
	end := Cell(endCol, endRow)
	return start + ":" + end
}

func splitCellRange(rangeRef string) (startCol, startRow, endCol, endRow int, err error) {
	parts := strings.SplitN(rangeRef, ":", 2)
	if len(parts) == 1 {
		startCol, startRow, err = CellNameToCoordinates(parts[0])
		if err != nil {
			return
		}
		endCol, endRow = startCol, startRow
		return
	}
	startCol, startRow, err = CellNameToCoordinates(parts[0])
	if err != nil {
		return
	}
	endCol, endRow, err = CellNameToCoordinates(parts[1])
	return
}
