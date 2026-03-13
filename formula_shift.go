package goods

import (
	"strconv"
	"strings"
	"unicode"
)

func shiftFormulaRefs(formula string, sheetName string, isRow bool, insertIdx int, delta int) string {
	var result strings.Builder
	i := 0

	for i < len(formula) {
		if formula[i] == '"' {
			result.WriteByte('"')
			i++
			for i < len(formula) && formula[i] != '"' {
				result.WriteByte(formula[i])
				i++
			}
			if i < len(formula) {
				result.WriteByte('"')
				i++
			}
			continue
		}

		if formula[i] != '[' {
			result.WriteByte(formula[i])
			i++
			continue
		}

		bracketStart := i
		i++

		refSheet := ""
		explicitSheet := false
		if i < len(formula) && formula[i] == '.' {
			refSheet = sheetName
			i++
		} else {
			explicitSheet = true
			nameStart := i
			for i < len(formula) && formula[i] != '.' && formula[i] != ']' {
				i++
			}
			refSheet = formula[nameStart:i]
			if i < len(formula) && formula[i] == '.' {
				i++
			} else {
				result.WriteString(formula[bracketStart : i+1])
				if i < len(formula) {
					i++
				}
				continue
			}
		}

		start1 := i
		for i < len(formula) && formula[i] != ']' && formula[i] != ':' {
			i++
		}
		ref1 := formula[start1:i]

		isRange := false
		ref2 := ""
		if i < len(formula) && formula[i] == ':' {
			isRange = true
			i++
			if i < len(formula) && formula[i] == '.' {
				i++
			}
			start2 := i
			for i < len(formula) && formula[i] != ']' {
				i++
			}
			ref2 = formula[start2:i]
		}
		if i < len(formula) {
			i++
		}

		if refSheet != sheetName {
			result.WriteString(formula[bracketStart:i])
			continue
		}

		newRef1 := shiftRef(ref1, isRow, insertIdx, delta)
		writeSheetPrefix := func() {
			result.WriteByte('[')
			if explicitSheet {
				result.WriteString(refSheet)
			}
			result.WriteByte('.')
		}
		if isRange {
			newRef2 := shiftRef(ref2, isRow, insertIdx, delta)
			if newRef1 == "#REF!" || newRef2 == "#REF!" {
				result.WriteString("#REF!")
			} else {
				writeSheetPrefix()
				result.WriteString(newRef1)
				result.WriteString(":.")
				result.WriteString(newRef2)
				result.WriteByte(']')
			}
		} else {
			if newRef1 == "#REF!" {
				result.WriteString("#REF!")
			} else {
				writeSheetPrefix()
				result.WriteString(newRef1)
				result.WriteByte(']')
			}
		}
	}

	return result.String()
}

func shiftRef(ref string, isRow bool, insertIdx int, delta int) string {
	colAbsolute := false
	rowAbsolute := false
	colStr := ""
	rowStr := ""

	j := 0
	if j < len(ref) && ref[j] == '$' {
		colAbsolute = true
		j++
	}
	for j < len(ref) && unicode.IsLetter(rune(ref[j])) {
		colStr += string(unicode.ToUpper(rune(ref[j])))
		j++
	}
	if j < len(ref) && ref[j] == '$' {
		rowAbsolute = true
		j++
	}
	for j < len(ref) && unicode.IsDigit(rune(ref[j])) {
		rowStr += string(ref[j])
		j++
	}

	if colStr == "" || rowStr == "" {
		return ref
	}

	col := columnNameToNumber(colStr)
	row, _ := strconv.Atoi(rowStr)

	if isRow {
		if rowAbsolute {
			return ref
		}
		if row >= insertIdx {
			if delta < 0 && row < insertIdx-delta {
				return "#REF!"
			}
			row += delta
			if row < 1 {
				return "#REF!"
			}
		}
	} else {
		if colAbsolute {
			return ref
		}
		if col >= insertIdx {
			if delta < 0 && col < insertIdx-delta {
				return "#REF!"
			}
			col += delta
			if col < 1 {
				return "#REF!"
			}
		}
		colStr = columnNumberToName(col)
	}

	var b strings.Builder
	if colAbsolute {
		b.WriteByte('$')
	}
	b.WriteString(colStr)
	if rowAbsolute {
		b.WriteByte('$')
	}
	b.WriteString(strconv.Itoa(row))
	return b.String()
}

func (f *File) shiftFormulasOnInsertRows(sheetName string, rowIdx, count int) {
	for _, s := range f.sheets {
		for _, r := range s.rows {
			for _, c := range r.cells {
				if c.formula != "" {
					c.formula = shiftFormulaRefs(c.formula, sheetName, true, rowIdx, count)
				}
			}
		}
	}
}

func (f *File) shiftFormulasOnRemoveRow(sheetName string, rowIdx int) {
	for _, s := range f.sheets {
		for _, r := range s.rows {
			for _, c := range r.cells {
				if c.formula != "" {
					c.formula = shiftFormulaRefs(c.formula, sheetName, true, rowIdx, -1)
				}
			}
		}
	}
}

func (f *File) shiftFormulasOnInsertCols(sheetName string, colIdx, count int) {
	for _, s := range f.sheets {
		for _, r := range s.rows {
			for _, c := range r.cells {
				if c.formula != "" {
					c.formula = shiftFormulaRefs(c.formula, sheetName, false, colIdx, count)
				}
			}
		}
	}
}

func (f *File) shiftFormulasOnRemoveCol(sheetName string, colIdx int) {
	for _, s := range f.sheets {
		for _, r := range s.rows {
			for _, c := range r.cells {
				if c.formula != "" {
					c.formula = shiftFormulaRefs(c.formula, sheetName, false, colIdx, -1)
				}
			}
		}
	}
}
