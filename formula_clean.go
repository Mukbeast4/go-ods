package goods

import (
	"regexp"
	"strconv"
	"strings"
)

var (
	reCellRef  = regexp.MustCompile(`\[\.([A-Z]+)(\d+)\]`)
	reRangeRef = regexp.MustCompile(`\[\.([A-Z]+)(\d+):\.([A-Z]+)(\d+)\]`)
)

// CleanFormula rewrites an ODS formula into a canonical form closer to common
// spreadsheet syntax: "[.A1:.B3]" becomes "A1:B3", "of:=" prefix is stripped,
// XML entities are decoded and ";" argument separators become ", ".
func CleanFormula(odsFormula string) string {
	f := strings.TrimPrefix(odsFormula, "of:=")
	f = strings.TrimPrefix(f, "of:")

	f = reRangeRef.ReplaceAllString(f, "$1$2:$3$4")
	f = reCellRef.ReplaceAllString(f, "$1$2")

	f = strings.ReplaceAll(f, "&gt;", ">")
	f = strings.ReplaceAll(f, "&lt;", "<")
	f = strings.ReplaceAll(f, "&amp;", "&")
	f = strings.ReplaceAll(f, "&quot;", `"`)
	f = strings.ReplaceAll(f, "&apos;", "'")

	f = strings.ReplaceAll(f, ";", ", ")

	return f
}

// AdaptFormula rewrites row indices that equal fromRow into toRow, preserving
// the ODS formula syntax. Use it to replicate a formula across rows.
func AdaptFormula(formula string, fromRow, toRow int) string {
	re := regexp.MustCompile(`(\.\s*[A-Z]+)` + strconv.Itoa(fromRow) + `([\]:])`)
	return re.ReplaceAllString(formula, "${1}"+strconv.Itoa(toRow)+"${2}")
}

// GetCleanFormula returns the cell formula run through CleanFormula. It
// returns "" when the cell has no formula.
func (f *File) GetCleanFormula(sheet, cellRef string) (string, error) {
	raw, err := f.GetCellFormula(sheet, cellRef)
	if err != nil {
		return "", err
	}
	if raw == "" {
		return "", nil
	}
	return CleanFormula(raw), nil
}

// GetSheetFormulas returns a snapshot of every formula on the sheet keyed by
// A1-style cell reference. Values are run through CleanFormula.
func (f *File) GetSheetFormulas(sheet string) (map[string]string, error) {
	if f.closed {
		return nil, ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
		return nil, ErrSheetNotFound
	}

	formulas := make(map[string]string)
	for rowIdx, r := range s.rows {
		for colIdx, c := range r.cells {
			if c.formula != "" {
				cellName, err := CoordinatesToCellName(colIdx, rowIdx)
				if err != nil {
					continue
				}
				formulas[cellName] = CleanFormula(c.formula)
			}
		}
	}

	return formulas, nil
}

// CopyRowFormulas copies every formula from fromRow into toRow, adapting row
// indices so references to fromRow become references to toRow.
func (f *File) CopyRowFormulas(sheet string, fromRow, toRow int) error {
	if f.closed {
		return ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
		return ErrSheetNotFound
	}

	srcRow, ok := s.rows[fromRow]
	if !ok {
		return ErrRowOutOfRange
	}

	dstRow := s.getOrCreateRow(toRow)

	for colIdx, c := range srcRow.cells {
		if c.formula == "" {
			continue
		}

		adapted := AdaptFormula(c.formula, fromRow, toRow)

		dstCell, ok := dstRow.cells[colIdx]
		if !ok {
			dstCell = &cell{valueType: c.valueType}
			dstRow.cells[colIdx] = dstCell
		}
		dstCell.formula = adapted

		if colIdx > s.maxCol {
			s.maxCol = colIdx
		}
	}

	if toRow > s.maxRow {
		s.maxRow = toRow
	}

	return nil
}
