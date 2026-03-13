package goods

import (
	"strconv"
	"strings"
	"unicode"
)

type bracketParseResult struct {
	refSheet      string
	ref1          string
	ref2          string
	isRange       bool
	explicitSheet bool
	endPos        int
	skip          bool
}

func parseShiftSheetName(formula string, pos int, defaultSheet string) (sheetName string, explicitSheet bool, newPos int, valid bool) {
	i := pos
	if i < len(formula) && formula[i] == '.' {
		return defaultSheet, false, i + 1, true
	}

	nameStart := i
	for i < len(formula) && formula[i] != '.' && formula[i] != ']' {
		i++
	}
	sheetName = formula[nameStart:i]
	if i < len(formula) && formula[i] == '.' {
		return sheetName, true, i + 1, true
	}
	if i < len(formula) {
		i++
	}
	return "", true, i, false
}

func parseShiftBracket(formula string, pos int, sheetName string) bracketParseResult {
	r := bracketParseResult{}

	refSheet, explicit, i, valid := parseShiftSheetName(formula, pos, sheetName)
	r.refSheet = refSheet
	r.explicitSheet = explicit
	if !valid {
		r.endPos = i
		r.skip = true
		return r
	}

	r.ref1, r.ref2, r.isRange, i = parseBracketRange(formula, i)
	r.endPos = i
	return r
}

func writeShiftedRef(result *strings.Builder, refSheet string, explicitSheet bool, ref1, ref2 string, isRange bool, isRow bool, insertIdx int, delta int) {
	newRef1 := shiftRef(ref1, isRow, insertIdx, delta)

	writePrefix := func() {
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
			writePrefix()
			result.WriteString(newRef1)
			result.WriteString(":.")
			result.WriteString(newRef2)
			result.WriteByte(']')
		}
	} else {
		if newRef1 == "#REF!" {
			result.WriteString("#REF!")
		} else {
			writePrefix()
			result.WriteString(newRef1)
			result.WriteByte(']')
		}
	}
}

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

		br := parseShiftBracket(formula, i, sheetName)
		i = br.endPos

		if br.skip {
			result.WriteString(formula[bracketStart:br.endPos])
			continue
		}

		if br.refSheet != sheetName {
			result.WriteString(formula[bracketStart:i])
			continue
		}

		writeShiftedRef(&result, br.refSheet, br.explicitSheet, br.ref1, br.ref2, br.isRange, isRow, insertIdx, delta)
	}

	return result.String()
}

func parseRefComponents(ref string) (colStr string, rowStr string, colAbsolute bool, rowAbsolute bool) {
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
	return
}

func applyShift(val int, absolute bool, insertIdx int, delta int) (int, bool) {
	if absolute {
		return val, false
	}
	if val >= insertIdx {
		if delta < 0 && val < insertIdx-delta {
			return 0, true
		}
		val += delta
		if val < 1 {
			return 0, true
		}
	}
	return val, false
}

func shiftRef(ref string, isRow bool, insertIdx int, delta int) string {
	colStr, rowStr, colAbsolute, rowAbsolute := parseRefComponents(ref)

	if colStr == "" || rowStr == "" {
		return ref
	}

	col := columnNameToNumber(colStr)
	row, _ := strconv.Atoi(rowStr)

	if isRow {
		newRow, isRefErr := applyShift(row, rowAbsolute, insertIdx, delta)
		if rowAbsolute {
			return ref
		}
		if isRefErr {
			return "#REF!"
		}
		row = newRow
	} else {
		newCol, isRefErr := applyShift(col, colAbsolute, insertIdx, delta)
		if colAbsolute {
			return ref
		}
		if isRefErr {
			return "#REF!"
		}
		col = newCol
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
