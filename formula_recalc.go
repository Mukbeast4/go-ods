package goods

import (
	"fmt"
	"strconv"
	"strings"
	"unicode"
)

type cellRef struct {
	sheet string
	col   int
	row   int
}

type depGraph struct {
	edges    map[cellRef][]cellRef
	reverse  map[cellRef][]cellRef
	formulas map[cellRef]string
}

func extractRefs(formula string) []cellRef {
	return extractRefsWithSheet(formula, "")
}

func parseBracketSheetName(formula string, pos int, currentSheet string) (sheetName string, newPos int, valid bool) {
	i := pos
	if i < len(formula) && formula[i] == '.' {
		return currentSheet, i + 1, true
	}

	nameStart := i
	for i < len(formula) && formula[i] != '.' && formula[i] != ']' {
		i++
	}
	sheetName = formula[nameStart:i]
	if i < len(formula) && formula[i] == '.' {
		return sheetName, i + 1, true
	}
	if i < len(formula) {
		i++
	}
	return "", i, false
}

func parseBracketRange(formula string, pos int) (ref1 string, ref2 string, isRange bool, newPos int) {
	i := pos
	start := i
	for i < len(formula) && formula[i] != ']' && formula[i] != ':' {
		i++
	}
	ref1 = formula[start:i]

	if i < len(formula) && formula[i] == ':' {
		i++
		if i < len(formula) && formula[i] == '.' {
			i++
		}
		start2 := i
		for i < len(formula) && formula[i] != ']' {
			i++
		}
		ref2 = formula[start2:i]
		if i < len(formula) {
			i++
		}
		return ref1, ref2, true, i
	}

	if i < len(formula) {
		i++
	}
	return ref1, "", false, i
}

func parseBracketRef(formula string, pos int, currentSheet string) ([]cellRef, int) {
	sheetName, i, valid := parseBracketSheetName(formula, pos, currentSheet)
	if !valid {
		return nil, i
	}

	ref1, ref2, isRange, i := parseBracketRange(formula, i)

	if isRange {
		rangeRefs := expandRange(ref1, ref2)
		for idx := range rangeRefs {
			rangeRefs[idx].sheet = sheetName
		}
		return rangeRefs, i
	}

	cr, ok := parseRef(ref1)
	if ok {
		cr.sheet = sheetName
		return []cellRef{cr}, i
	}
	return nil, i
}

func extractRefsWithSheet(formula string, currentSheet string) []cellRef {
	var refs []cellRef
	i := 0
	for i < len(formula) {
		if formula[i] == '"' {
			i++
			for i < len(formula) && formula[i] != '"' {
				i++
			}
			if i < len(formula) {
				i++
			}
			continue
		}

		if formula[i] != '[' {
			i++
			continue
		}

		i++
		found, newPos := parseBracketRef(formula, i, currentSheet)
		i = newPos
		refs = append(refs, found...)
	}
	return refs
}

func parseRef(ref string) (cellRef, bool) {
	name := extractCellName(ref)
	if name == "" {
		return cellRef{}, false
	}
	col, row, err := CellNameToCoordinates(name)
	if err != nil {
		return cellRef{}, false
	}
	return cellRef{col: col, row: row}, true
}

func expandRange(startRef, endRef string) []cellRef {
	startCol, startRow := splitRef(startRef)
	endCol, endRow := splitRef(endRef)

	sc := columnNameToNumber(startCol)
	ec := columnNameToNumber(endCol)

	sr, _ := strconv.Atoi(startRow)
	er, _ := strconv.Atoi(endRow)

	if sr == 0 || er == 0 || sc == 0 || ec == 0 {
		return nil
	}

	var refs []cellRef
	for r := sr; r <= er; r++ {
		for c := sc; c <= ec; c++ {
			refs = append(refs, cellRef{col: c, row: r})
		}
	}
	return refs
}

func buildDepGraph(s *sheet) *depGraph {
	return buildDepGraphForSheet(s, "")
}

func buildDepGraphForSheet(s *sheet, sheetName string) *depGraph {
	g := &depGraph{
		edges:    make(map[cellRef][]cellRef),
		reverse:  make(map[cellRef][]cellRef),
		formulas: make(map[cellRef]string),
	}

	for rowIdx, r := range s.rows {
		for colIdx, c := range r.cells {
			if c.formula == "" {
				continue
			}

			cr := cellRef{sheet: sheetName, col: colIdx, row: rowIdx}
			g.formulas[cr] = c.formula

			deps := extractRefsWithSheet(c.formula, sheetName)
			g.edges[cr] = deps

			for _, dep := range deps {
				g.reverse[dep] = append(g.reverse[dep], cr)
			}
		}
	}

	return g
}

func (g *depGraph) topoSort() ([]cellRef, error) {
	inDegree := make(map[cellRef]int)

	for cr := range g.formulas {
		if _, ok := inDegree[cr]; !ok {
			inDegree[cr] = 0
		}
	}

	for cr, deps := range g.edges {
		count := 0
		for _, dep := range deps {
			if _, isFormula := g.formulas[dep]; isFormula {
				count++
			}
		}
		inDegree[cr] = count
	}

	var queue []cellRef
	for cr, deg := range inDegree {
		if deg == 0 {
			queue = append(queue, cr)
		}
	}

	var sorted []cellRef
	for len(queue) > 0 {
		cr := queue[0]
		queue = queue[1:]
		sorted = append(sorted, cr)

		for _, dependent := range g.reverse[cr] {
			if _, isFormula := g.formulas[dependent]; !isFormula {
				continue
			}
			inDegree[dependent]--
			if inDegree[dependent] == 0 {
				queue = append(queue, dependent)
			}
		}
	}

	if len(sorted) != len(g.formulas) {
		return nil, ErrCircularReference
	}

	return sorted, nil
}

func collectCellValues(s *sheet) CellValues {
	values := make(CellValues)
	for rowIdx, r := range s.rows {
		for colIdx, c := range r.cells {
			if c.rawValue == "" {
				continue
			}
			name, err := CoordinatesToCellName(colIdx, rowIdx)
			if err != nil {
				continue
			}
			if f, fErr := strconv.ParseFloat(c.rawValue, 64); fErr == nil {
				values[name] = f
			} else if c.valueType == CellTypeBool {
				values[name] = c.rawValue == "true"
			} else {
				values[name] = c.rawValue
			}
		}
	}
	return values
}

func storeResult(c *cell, result interface{}) {
	switch v := result.(type) {
	case float64:
		c.valueType = CellTypeFloat
		c.rawValue = strconv.FormatFloat(v, 'f', -1, 64)
	case bool:
		c.valueType = CellTypeBool
		if v {
			c.rawValue = "true"
		} else {
			c.rawValue = "false"
		}
	case string:
		c.valueType = CellTypeString
		c.rawValue = v
	default:
		c.valueType = CellTypeString
		c.rawValue = fmt.Sprintf("%v", v)
	}
}

func (f *File) RecalcSheet(sheetName string) error {
	if f.closed {
		return ErrFileClosed
	}
	s := f.getSheet(sheetName)
	if s == nil {
		return ErrSheetNotFound
	}

	g := buildDepGraph(s)
	if len(g.formulas) == 0 {
		return nil
	}

	sorted, err := g.topoSort()
	if err != nil {
		return err
	}

	values := collectCellValues(s)

	for _, cr := range sorted {
		formula := g.formulas[cr]

		p := &formulaParser{input: formula, pos: 0, values: values, file: f}
		result, err := p.parseExpr()
		if err != nil {
			continue
		}

		c := s.getCell(cr.col, cr.row)
		if c == nil {
			continue
		}
		storeResult(c, result)

		name, nameErr := CoordinatesToCellName(cr.col, cr.row)
		if nameErr == nil {
			values[name] = result
		}
	}

	return nil
}

func (f *File) RecalcAll() error {
	if f.closed {
		return ErrFileClosed
	}

	g := &depGraph{
		edges:    make(map[cellRef][]cellRef),
		reverse:  make(map[cellRef][]cellRef),
		formulas: make(map[cellRef]string),
	}

	for _, s := range f.sheets {
		sg := buildDepGraphForSheet(s, s.name)
		for cr, formula := range sg.formulas {
			g.formulas[cr] = formula
		}
		for cr, deps := range sg.edges {
			g.edges[cr] = deps
		}
		for dep, dependents := range sg.reverse {
			g.reverse[dep] = append(g.reverse[dep], dependents...)
		}
	}

	if len(g.formulas) == 0 {
		return nil
	}

	sorted, err := g.topoSort()
	if err != nil {
		return err
	}

	allValues := make(map[string]CellValues)
	for _, s := range f.sheets {
		allValues[s.name] = collectCellValues(s)
	}

	for _, cr := range sorted {
		formula := g.formulas[cr]
		s := f.getSheet(cr.sheet)
		if s == nil {
			continue
		}

		values := allValues[cr.sheet]
		p := &formulaParser{input: formula, pos: 0, values: values, file: f}
		result, err := p.parseExpr()
		if err != nil {
			continue
		}

		c := s.getCell(cr.col, cr.row)
		if c == nil {
			continue
		}
		storeResult(c, result)

		name, nameErr := CoordinatesToCellName(cr.col, cr.row)
		if nameErr == nil {
			values[name] = result
		}
	}

	return nil
}

func (f *File) SetAutoRecalc(enabled bool) {
	f.autoRecalc = enabled
}

func (f *File) triggerRecalc(sheetName string) {
	if !f.autoRecalc {
		return
	}
	_ = f.RecalcSheet(sheetName)
}

func hasFormulaRefs(s *sheet) bool {
	for _, r := range s.rows {
		for _, c := range r.cells {
			if c.formula != "" {
				return true
			}
		}
	}
	return false
}

func stripDollarSigns(ref string) string {
	var b strings.Builder
	for _, c := range ref {
		if c != '$' {
			b.WriteRune(c)
		}
	}
	return b.String()
}

func extractCellNameClean(ref string) string {
	ref = stripDollarSigns(ref)
	var result strings.Builder
	for _, c := range ref {
		if unicode.IsLetter(c) || unicode.IsDigit(c) {
			result.WriteRune(unicode.ToUpper(c))
		}
	}
	return result.String()
}
