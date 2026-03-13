package goods

import (
	"fmt"
	"math"
	"strconv"
	"strings"
	"time"
	"unicode"
)

type CellValues map[string]interface{}

type rangeData struct {
	values []interface{}
	rows   int
	cols   int
}

func Evaluate(formula string, values CellValues) (interface{}, error) {
	f := strings.TrimPrefix(formula, "of:=")
	f = strings.TrimPrefix(f, "of:")

	p := &formulaParser{input: f, pos: 0, values: values}
	result, err := p.parseExpr()
	if err != nil {
		return nil, fmt.Errorf("eval %q: %w", formula, err)
	}
	return result, nil
}

func (f *File) EvaluateFormula(sheet, cellRef string, extraValues CellValues) (interface{}, error) {
	if f.closed {
		return nil, ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
		return nil, ErrSheetNotFound
	}

	col, rowIdx, err := CellNameToCoordinates(cellRef)
	if err != nil {
		return nil, err
	}

	c := s.getCell(col, rowIdx)
	if c == nil || c.formula == "" {
		return nil, fmt.Errorf("goods: no formula at %s", cellRef)
	}

	values := collectCellValues(s)
	if extraValues != nil {
		for k, v := range extraValues {
			values[k] = v
		}
	}

	p := &formulaParser{input: c.formula, pos: 0, values: values, file: f}
	result, err := p.parseExpr()
	if err != nil {
		return nil, fmt.Errorf("eval %q: %w", c.formula, err)
	}
	return result, nil
}

type formulaParser struct {
	input  string
	pos    int
	values CellValues
	file   *File
}

func (p *formulaParser) peek() byte {
	p.skipSpaces()
	if p.pos >= len(p.input) {
		return 0
	}
	return p.input[p.pos]
}

func (p *formulaParser) skipSpaces() {
	for p.pos < len(p.input) && p.input[p.pos] == ' ' {
		p.pos++
	}
}

func (p *formulaParser) parseExpr() (interface{}, error) {
	return p.parseComparison()
}

func (p *formulaParser) parseComparison() (interface{}, error) {
	left, err := p.parseAddSub()
	if err != nil {
		return nil, err
	}

	for {
		p.skipSpaces()
		if p.pos >= len(p.input) {
			break
		}

		var op string
		if p.pos+1 < len(p.input) && p.input[p.pos:p.pos+2] == "<>" {
			op = "<>"
			p.pos += 2
		} else if p.pos+1 < len(p.input) && p.input[p.pos:p.pos+2] == "<=" {
			op = "<="
			p.pos += 2
		} else if p.pos+1 < len(p.input) && p.input[p.pos:p.pos+2] == ">=" {
			op = ">="
			p.pos += 2
		} else if p.input[p.pos] == '=' {
			op = "="
			p.pos++
		} else if p.input[p.pos] == '<' {
			op = "<"
			p.pos++
		} else if p.input[p.pos] == '>' {
			op = ">"
			p.pos++
		} else {
			break
		}

		right, err := p.parseAddSub()
		if err != nil {
			return nil, err
		}

		left = evalComparison(left, op, right)
	}

	return left, nil
}

func (p *formulaParser) parseAddSub() (interface{}, error) {
	left, err := p.parseMulDiv()
	if err != nil {
		return nil, err
	}

	for {
		p.skipSpaces()
		if p.pos >= len(p.input) {
			break
		}
		ch := p.input[p.pos]
		if ch != '+' && ch != '-' {
			break
		}

		if ch == '-' {
			if p.pos+1 < len(p.input) && (unicode.IsDigit(rune(p.input[p.pos+1])) || p.input[p.pos+1] == '.') {
				if _, isNum := toFloat(left); !isNum {
					break
				}
			}
		}

		p.pos++
		right, err := p.parseMulDiv()
		if err != nil {
			return nil, err
		}
		left, err = applyAddSub(left, right, ch)
		if err != nil {
			return nil, err
		}
	}

	return left, nil
}

func applyAddSub(left, right interface{}, op byte) (interface{}, error) {
	lf, lok := toFloat(left)
	rf, rok := toFloat(right)
	if lok && rok {
		if op == '+' {
			return lf + rf, nil
		}
		return lf - rf, nil
	}
	if op == '+' {
		return fmt.Sprintf("%v%v", left, right), nil
	}
	return nil, fmt.Errorf("cannot subtract non-numeric values")
}

func (p *formulaParser) parseMulDiv() (interface{}, error) {
	left, err := p.parseUnary()
	if err != nil {
		return nil, err
	}

	for {
		p.skipSpaces()
		if p.pos >= len(p.input) {
			break
		}
		ch := p.input[p.pos]
		if ch != '*' && ch != '/' {
			break
		}

		p.pos++
		right, err := p.parseUnary()
		if err != nil {
			return nil, err
		}
		lf, lok := toFloat(left)
		rf, rok := toFloat(right)
		if !lok || !rok {
			return nil, fmt.Errorf("cannot multiply/divide non-numeric values")
		}
		if ch == '*' {
			left = lf * rf
		} else {
			if rf == 0 {
				return nil, fmt.Errorf("division by zero")
			}
			left = lf / rf
		}
	}

	return left, nil
}

func (p *formulaParser) parseUnary() (interface{}, error) {
	p.skipSpaces()
	if p.pos < len(p.input) && p.input[p.pos] == '-' {
		p.pos++
		val, err := p.parsePrimary()
		if err != nil {
			return nil, err
		}
		if f, ok := toFloat(val); ok {
			return -f, nil
		}
		return nil, fmt.Errorf("cannot negate non-numeric value: %v", val)
	}
	if p.pos < len(p.input) && p.input[p.pos] == '+' {
		p.pos++
	}
	return p.parsePrimary()
}

func (p *formulaParser) parsePrimary() (interface{}, error) {
	p.skipSpaces()
	if p.pos >= len(p.input) {
		return nil, fmt.Errorf("unexpected end of expression")
	}

	ch := p.input[p.pos]

	if ch == '(' {
		p.pos++
		val, err := p.parseExpr()
		if err != nil {
			return nil, err
		}
		p.skipSpaces()
		if p.pos < len(p.input) && p.input[p.pos] == ')' {
			p.pos++
		}
		return val, nil
	}

	if ch == '"' {
		return p.parseStringLiteral()
	}

	if ch == '[' {
		return p.parseCellReference()
	}

	if unicode.IsDigit(rune(ch)) || ch == '.' {
		return p.parseNumberLiteral()
	}

	if unicode.IsLetter(rune(ch)) {
		return p.parseFunctionCall()
	}

	return nil, fmt.Errorf("unexpected character %q at pos %d", string(ch), p.pos)
}

func (p *formulaParser) parseStringLiteral() (interface{}, error) {
	p.pos++
	start := p.pos
	for p.pos < len(p.input) && p.input[p.pos] != '"' {
		p.pos++
	}
	val := p.input[start:p.pos]
	if p.pos < len(p.input) {
		p.pos++
	}
	return val, nil
}

func (p *formulaParser) parseNumberLiteral() (interface{}, error) {
	start := p.pos
	for p.pos < len(p.input) && (unicode.IsDigit(rune(p.input[p.pos])) || p.input[p.pos] == '.') {
		p.pos++
	}
	val, err := strconv.ParseFloat(p.input[start:p.pos], 64)
	if err != nil {
		return nil, fmt.Errorf("invalid number %q", p.input[start:p.pos])
	}
	return val, nil
}

func (p *formulaParser) parseSheetFromBracket() string {
	if p.pos < len(p.input) && p.input[p.pos] == '.' {
		p.pos++
		return ""
	}
	nameStart := p.pos
	for p.pos < len(p.input) && p.input[p.pos] != '.' && p.input[p.pos] != ']' {
		p.pos++
	}
	sheetName := p.input[nameStart:p.pos]
	if p.pos < len(p.input) && p.input[p.pos] == '.' {
		p.pos++
	}
	return sheetName
}

func (p *formulaParser) parseCellReference() (interface{}, error) {
	p.pos++

	sheetName := p.parseSheetFromBracket()

	start := p.pos
	for p.pos < len(p.input) && p.input[p.pos] != ']' && p.input[p.pos] != ':' {
		p.pos++
	}
	ref := p.input[start:p.pos]

	if p.pos < len(p.input) && p.input[p.pos] == ':' {
		p.pos++
		if p.pos < len(p.input) && p.input[p.pos] == '.' {
			p.pos++
		}
		endStart := p.pos
		for p.pos < len(p.input) && p.input[p.pos] != ']' {
			p.pos++
		}
		endRef := p.input[endStart:p.pos]
		if p.pos < len(p.input) {
			p.pos++
		}
		if sheetName != "" {
			return p.evalRangeFromSheet(sheetName, ref, endRef)
		}
		return p.evalRange(ref, endRef)
	}

	if p.pos < len(p.input) {
		p.pos++
	}

	if sheetName != "" {
		return p.evalCrossSheetRef(sheetName, ref)
	}

	cellName := extractCellName(ref)
	val, ok := p.values[cellName]
	if !ok {
		return "", nil
	}
	return val, nil
}

func (p *formulaParser) evalCrossSheetRef(sheetName, ref string) (interface{}, error) {
	if p.file == nil {
		return "", nil
	}
	s := p.file.getSheet(sheetName)
	if s == nil {
		return "", nil
	}
	cellName := extractCellName(ref)
	col, rowIdx, err := CellNameToCoordinates(cellName)
	if err != nil {
		return "", nil
	}
	c := s.getCell(col, rowIdx)
	if c == nil {
		return "", nil
	}
	if f, fErr := strconv.ParseFloat(c.rawValue, 64); fErr == nil {
		return f, nil
	}
	if c.valueType == CellTypeBool {
		return c.rawValue == "true", nil
	}
	return c.rawValue, nil
}

func (p *formulaParser) evalRangeFromSheet(sheetName, startRef, endRef string) (interface{}, error) {
	if p.file == nil {
		return &rangeData{}, nil
	}
	s := p.file.getSheet(sheetName)
	if s == nil {
		return &rangeData{}, nil
	}
	values := collectCellValues(s)
	tmp := &formulaParser{values: values}
	return tmp.evalRange(startRef, endRef)
}

func (p *formulaParser) evalRange(startRef, endRef string) (interface{}, error) {
	startCol, startRow := splitRef(startRef)
	endCol, endRow := splitRef(endRef)

	startColIdx := columnNameToNumber(startCol)
	endColIdx := columnNameToNumber(endCol)

	startRowNum := 0
	endRowNum := 0
	if startRow != "" {
		startRowNum, _ = strconv.Atoi(startRow)
	}
	if endRow != "" {
		endRowNum, _ = strconv.Atoi(endRow)
	}

	numCols := endColIdx - startColIdx + 1
	numRows := 1
	var vals []interface{}

	if startRowNum > 0 && endRowNum > 0 && startRowNum != endRowNum {
		numRows = endRowNum - startRowNum + 1
		for ri := startRowNum; ri <= endRowNum; ri++ {
			for ci := startColIdx; ci <= endColIdx; ci++ {
				name := columnNumberToName(ci) + strconv.Itoa(ri)
				if v, ok := p.values[name]; ok {
					vals = append(vals, v)
				} else {
					vals = append(vals, "")
				}
			}
		}
	} else {
		rowStr := startRow
		for ci := startColIdx; ci <= endColIdx; ci++ {
			name := columnNumberToName(ci) + rowStr
			if v, ok := p.values[name]; ok {
				vals = append(vals, v)
			} else {
				vals = append(vals, "")
			}
		}
	}

	return &rangeData{values: vals, rows: numRows, cols: numCols}, nil
}

func (p *formulaParser) parseFunctionCall() (interface{}, error) {
	start := p.pos
	for p.pos < len(p.input) && (unicode.IsLetter(rune(p.input[p.pos])) || unicode.IsDigit(rune(p.input[p.pos])) || p.input[p.pos] == '_') {
		p.pos++
	}
	name := strings.ToUpper(p.input[start:p.pos])

	p.skipSpaces()
	if p.pos >= len(p.input) || p.input[p.pos] != '(' {
		return nil, fmt.Errorf("expected '(' after function %s", name)
	}
	p.pos++

	if name == "IFERROR" {
		return p.parseIFERROR()
	}

	args, err := p.parseFunctionArgs()
	if err != nil {
		return nil, err
	}

	return evalFunction(name, args)
}

func (p *formulaParser) parseFunctionArgs() ([]interface{}, error) {
	var args []interface{}
	for {
		p.skipSpaces()
		if p.pos < len(p.input) && p.input[p.pos] == ')' {
			p.pos++
			break
		}

		arg, err := p.parseExpr()
		if err != nil {
			return nil, err
		}
		args = append(args, arg)

		p.skipSpaces()
		if p.pos < len(p.input) && p.input[p.pos] == ';' {
			p.pos++
		} else if p.pos < len(p.input) && p.input[p.pos] == ',' {
			p.pos++
		} else if p.pos < len(p.input) && p.input[p.pos] == ')' {
			p.pos++
			break
		}
	}
	return args, nil
}

func (p *formulaParser) parseIFERROR() (interface{}, error) {
	p.skipSpaces()
	val, err := p.parseExpr()

	p.skipSpaces()
	if p.pos < len(p.input) && (p.input[p.pos] == ';' || p.input[p.pos] == ',') {
		p.pos++
	}

	p.skipSpaces()
	fallback, fallbackErr := p.parseExpr()

	p.skipSpaces()
	if p.pos < len(p.input) && p.input[p.pos] == ')' {
		p.pos++
	}

	if err != nil {
		if fallbackErr != nil {
			return nil, fallbackErr
		}
		return fallback, nil
	}
	return val, nil
}

func isEmptyValue(v interface{}) bool {
	s, ok := v.(string)
	return ok && s == ""
}

func flattenArg(a interface{}) []interface{} {
	switch v := a.(type) {
	case []interface{}:
		return v
	case *rangeData:
		return v.values
	}
	return nil
}

func parseCriteriaOperator(criteria string) (string, string) {
	if len(criteria) >= 2 {
		switch criteria[:2] {
		case ">=", "<=", "<>":
			return criteria[:2], criteria[2:]
		}
	}
	if len(criteria) >= 1 {
		switch criteria[0] {
		case '>', '<', '=':
			return string(criteria[0]), criteria[1:]
		}
	}
	return "", criteria
}

func matchesCriteria(val interface{}, criteria string) bool {
	op, value := parseCriteriaOperator(criteria)

	if op == "<>" {
		return fmt.Sprintf("%v", val) != value
	}
	if op == "=" {
		return fmt.Sprintf("%v", val) == value
	}
	if op != "" {
		f, ok := toFloat(val)
		cf, cok := strconv.ParseFloat(value, 64)
		if !ok || cok != nil {
			return false
		}
		switch op {
		case ">=":
			return f >= cf
		case "<=":
			return f <= cf
		case ">":
			return f > cf
		case "<":
			return f < cf
		}
	}

	if cf, err := strconv.ParseFloat(criteria, 64); err == nil {
		f, ok := toFloat(val)
		return ok && f == cf
	}
	return strings.EqualFold(fmt.Sprintf("%v", val), criteria)
}

type formulaFunc func([]interface{}) (interface{}, error)

var formulaFunctions = map[string]formulaFunc{
	"IF":          evalIF,
	"AND":         evalAND,
	"OR":          evalOR,
	"NOT":         evalNOT,
	"SUM":         evalSUM,
	"AVERAGE":     evalAVERAGE,
	"COUNT":       evalCOUNT,
	"ABS":         evalABS,
	"MIN":         evalMIN,
	"MAX":         evalMAX,
	"ROUND":       evalROUND,
	"FLOOR":       evalFLOOR,
	"CEIL":        evalCEILING,
	"CEILING":     evalCEILING,
	"CONCATENATE": evalCONCATENATE,
	"LEN":         evalLEN,
	"TRIM":        evalTRIM,
	"UPPER":       evalUPPER,
	"LOWER":       evalLOWER,
	"LEFT":        evalLEFT,
	"RIGHT":       evalRIGHT,
	"MID":         evalMID,
	"MOD":         evalMOD,
	"POWER":       evalPOWER,
	"SQRT":        evalSQRT,
	"INT":         evalINT,
	"COUNTA":      evalCOUNTA,
	"SUMIF":       evalSUMIF,
	"COUNTIF":     evalCOUNTIF,
	"COUNTIFS":    evalCOUNTIFS,
	"SUMIFS":      evalSUMIFS,
	"VLOOKUP":     evalVLOOKUP,
	"HLOOKUP":     evalHLOOKUP,
	"INDEX":       evalINDEX,
	"MATCH":       evalMATCH,
	"DATE":        evalDATE,
	"TODAY":       evalTODAY,
	"NOW":         evalNOW,
	"YEAR":        evalYEAR,
	"MONTH":       evalMONTH,
	"DAY":         evalDAY,
	"FIND":        evalFIND,
	"SEARCH":      evalSEARCH,
	"SUBSTITUTE":  evalSUBSTITUTE,
	"REPLACE":     evalREPLACE,
	"TEXT":        evalTEXT,
	"VALUE":       evalVALUE,
	"SUMPRODUCT":  evalSUMPRODUCT,
}

func evalFunction(name string, args []interface{}) (interface{}, error) {
	if fn, ok := formulaFunctions[name]; ok {
		return fn(args)
	}
	return nil, fmt.Errorf("unknown function: %s", name)
}

func evalIF(args []interface{}) (interface{}, error) {
	if len(args) < 2 {
		return nil, fmt.Errorf("IF requires at least 2 arguments")
	}
	if toBool(args[0]) {
		return args[1], nil
	}
	if len(args) >= 3 {
		return args[2], nil
	}
	return false, nil
}

func evalAND(args []interface{}) (interface{}, error) {
	for _, a := range args {
		if !toBool(a) {
			return false, nil
		}
	}
	return true, nil
}

func evalOR(args []interface{}) (interface{}, error) {
	for _, a := range args {
		if toBool(a) {
			return true, nil
		}
	}
	return false, nil
}

func evalNOT(args []interface{}) (interface{}, error) {
	if len(args) != 1 {
		return nil, fmt.Errorf("NOT requires 1 argument")
	}
	return !toBool(args[0]), nil
}

func evalSUM(args []interface{}) (interface{}, error) {
	sum := 0.0
	for _, a := range args {
		if vals := flattenArg(a); vals != nil {
			for _, v := range vals {
				if f, ok := toFloat(v); ok {
					sum += f
				}
			}
		} else if f, ok := toFloat(a); ok {
			sum += f
		}
	}
	return sum, nil
}

func evalAVERAGE(args []interface{}) (interface{}, error) {
	sum := 0.0
	count := 0
	for _, a := range args {
		if vals := flattenArg(a); vals != nil {
			for _, v := range vals {
				if isEmptyValue(v) {
					continue
				}
				if f, ok := toFloat(v); ok {
					sum += f
					count++
				}
			}
		} else {
			if isEmptyValue(a) {
				continue
			}
			if f, ok := toFloat(a); ok {
				sum += f
				count++
			}
		}
	}
	if count == 0 {
		return nil, fmt.Errorf("AVERAGE: no numeric values")
	}
	return sum / float64(count), nil
}

func evalCOUNT(args []interface{}) (interface{}, error) {
	count := 0
	for _, a := range args {
		if vals := flattenArg(a); vals != nil {
			for _, v := range vals {
				if isEmptyValue(v) {
					continue
				}
				if _, ok := toFloat(v); ok {
					count++
				}
			}
		} else {
			if isEmptyValue(a) {
				continue
			}
			if _, ok := toFloat(a); ok {
				count++
			}
		}
	}
	return float64(count), nil
}

func evalABS(args []interface{}) (interface{}, error) {
	if len(args) != 1 {
		return nil, fmt.Errorf("ABS requires 1 argument")
	}
	if f, ok := toFloat(args[0]); ok {
		return math.Abs(f), nil
	}
	return nil, fmt.Errorf("ABS requires numeric argument")
}

func evalMIN(args []interface{}) (interface{}, error) {
	minVal := math.Inf(1)
	found := false
	for _, a := range args {
		if vals := flattenArg(a); vals != nil {
			for _, v := range vals {
				if f, ok := toFloat(v); ok {
					found = true
					if f < minVal {
						minVal = f
					}
				}
			}
		} else if f, ok := toFloat(a); ok {
			found = true
			if f < minVal {
				minVal = f
			}
		}
	}
	if !found {
		return float64(0), nil
	}
	return minVal, nil
}

func evalMAX(args []interface{}) (interface{}, error) {
	maxVal := math.Inf(-1)
	found := false
	for _, a := range args {
		if vals := flattenArg(a); vals != nil {
			for _, v := range vals {
				if f, ok := toFloat(v); ok {
					found = true
					if f > maxVal {
						maxVal = f
					}
				}
			}
		} else if f, ok := toFloat(a); ok {
			found = true
			if f > maxVal {
				maxVal = f
			}
		}
	}
	if !found {
		return float64(0), nil
	}
	return maxVal, nil
}

func evalROUND(args []interface{}) (interface{}, error) {
	if len(args) < 1 {
		return nil, fmt.Errorf("ROUND requires at least 1 argument")
	}
	f, ok := toFloat(args[0])
	if !ok {
		return nil, fmt.Errorf("ROUND requires numeric argument")
	}
	decimals := 0
	if len(args) >= 2 {
		if d, ok := toFloat(args[1]); ok {
			decimals = int(d)
		}
	}
	pow := math.Pow(10, float64(decimals))
	return math.Round(f*pow) / pow, nil
}

func evalFLOOR(args []interface{}) (interface{}, error) {
	if len(args) < 1 {
		return nil, fmt.Errorf("FLOOR requires at least 1 argument")
	}
	f, ok := toFloat(args[0])
	if !ok {
		return nil, fmt.Errorf("FLOOR requires numeric argument")
	}
	return math.Floor(f), nil
}

func evalCEILING(args []interface{}) (interface{}, error) {
	if len(args) < 1 {
		return nil, fmt.Errorf("CEILING requires at least 1 argument")
	}
	f, ok := toFloat(args[0])
	if !ok {
		return nil, fmt.Errorf("CEILING requires numeric argument")
	}
	return math.Ceil(f), nil
}

func evalCONCATENATE(args []interface{}) (interface{}, error) {
	var sb strings.Builder
	for _, a := range args {
		sb.WriteString(fmt.Sprintf("%v", a))
	}
	return sb.String(), nil
}

func evalLEN(args []interface{}) (interface{}, error) {
	if len(args) != 1 {
		return nil, fmt.Errorf("LEN requires 1 argument")
	}
	return float64(len(fmt.Sprintf("%v", args[0]))), nil
}

func evalTRIM(args []interface{}) (interface{}, error) {
	if len(args) != 1 {
		return nil, fmt.Errorf("TRIM requires 1 argument")
	}
	return strings.TrimSpace(fmt.Sprintf("%v", args[0])), nil
}

func evalUPPER(args []interface{}) (interface{}, error) {
	if len(args) != 1 {
		return nil, fmt.Errorf("UPPER requires 1 argument")
	}
	return strings.ToUpper(fmt.Sprintf("%v", args[0])), nil
}

func evalLOWER(args []interface{}) (interface{}, error) {
	if len(args) != 1 {
		return nil, fmt.Errorf("LOWER requires 1 argument")
	}
	return strings.ToLower(fmt.Sprintf("%v", args[0])), nil
}

func evalLEFT(args []interface{}) (interface{}, error) {
	if len(args) != 2 {
		return nil, fmt.Errorf("LEFT requires 2 arguments")
	}
	s := fmt.Sprintf("%v", args[0])
	n, ok := toFloat(args[1])
	if !ok {
		return nil, fmt.Errorf("LEFT requires numeric second argument")
	}
	idx := int(n)
	if idx > len(s) {
		idx = len(s)
	}
	return s[:idx], nil
}

func evalRIGHT(args []interface{}) (interface{}, error) {
	if len(args) != 2 {
		return nil, fmt.Errorf("RIGHT requires 2 arguments")
	}
	s := fmt.Sprintf("%v", args[0])
	n, ok := toFloat(args[1])
	if !ok {
		return nil, fmt.Errorf("RIGHT requires numeric second argument")
	}
	idx := int(n)
	if idx > len(s) {
		idx = len(s)
	}
	return s[len(s)-idx:], nil
}

func evalMID(args []interface{}) (interface{}, error) {
	if len(args) != 3 {
		return nil, fmt.Errorf("MID requires 3 arguments")
	}
	s := fmt.Sprintf("%v", args[0])
	start, ok1 := toFloat(args[1])
	length, ok2 := toFloat(args[2])
	if !ok1 || !ok2 {
		return nil, fmt.Errorf("MID requires numeric arguments")
	}
	startIdx := int(start) - 1
	if startIdx < 0 {
		startIdx = 0
	}
	endIdx := startIdx + int(length)
	if endIdx > len(s) {
		endIdx = len(s)
	}
	if startIdx >= len(s) {
		return "", nil
	}
	return s[startIdx:endIdx], nil
}

func evalMOD(args []interface{}) (interface{}, error) {
	if len(args) != 2 {
		return nil, fmt.Errorf("MOD requires 2 arguments")
	}
	a, ok1 := toFloat(args[0])
	b, ok2 := toFloat(args[1])
	if !ok1 || !ok2 {
		return nil, fmt.Errorf("MOD requires numeric arguments")
	}
	if b == 0 {
		return nil, fmt.Errorf("MOD: division by zero")
	}
	return math.Mod(a, b), nil
}

func evalPOWER(args []interface{}) (interface{}, error) {
	if len(args) != 2 {
		return nil, fmt.Errorf("POWER requires 2 arguments")
	}
	base, ok1 := toFloat(args[0])
	exp, ok2 := toFloat(args[1])
	if !ok1 || !ok2 {
		return nil, fmt.Errorf("POWER requires numeric arguments")
	}
	return math.Pow(base, exp), nil
}

func evalSQRT(args []interface{}) (interface{}, error) {
	if len(args) != 1 {
		return nil, fmt.Errorf("SQRT requires 1 argument")
	}
	f, ok := toFloat(args[0])
	if !ok {
		return nil, fmt.Errorf("SQRT requires numeric argument")
	}
	return math.Sqrt(f), nil
}

func evalINT(args []interface{}) (interface{}, error) {
	if len(args) != 1 {
		return nil, fmt.Errorf("INT requires 1 argument")
	}
	f, ok := toFloat(args[0])
	if !ok {
		return nil, fmt.Errorf("INT requires numeric argument")
	}
	return math.Floor(f), nil
}

func evalCOUNTA(args []interface{}) (interface{}, error) {
	count := 0
	for _, a := range args {
		if vals := flattenArg(a); vals != nil {
			for _, v := range vals {
				if v != nil {
					s := fmt.Sprintf("%v", v)
					if s != "" {
						count++
					}
				}
			}
		} else if a != nil {
			s := fmt.Sprintf("%v", a)
			if s != "" {
				count++
			}
		}
	}
	return float64(count), nil
}

func evalSUMIF(args []interface{}) (interface{}, error) {
	if len(args) < 2 {
		return nil, fmt.Errorf("SUMIF requires at least 2 arguments")
	}
	rangeVals := flattenArg(args[0])
	if rangeVals == nil {
		return float64(0), nil
	}
	criteria := fmt.Sprintf("%v", args[1])
	sumVals := rangeVals
	if len(args) >= 3 {
		sumVals = flattenArg(args[2])
		if sumVals == nil {
			return float64(0), nil
		}
	}
	sum := 0.0
	for i, v := range rangeVals {
		if matchesCriteria(v, criteria) && i < len(sumVals) {
			if f, ok := toFloat(sumVals[i]); ok {
				sum += f
			}
		}
	}
	return sum, nil
}

func evalCOUNTIF(args []interface{}) (interface{}, error) {
	if len(args) != 2 {
		return nil, fmt.Errorf("COUNTIF requires 2 arguments")
	}
	rangeVals := flattenArg(args[0])
	if rangeVals == nil {
		return float64(0), nil
	}
	criteria := fmt.Sprintf("%v", args[1])
	count := 0
	for _, v := range rangeVals {
		if matchesCriteria(v, criteria) {
			count++
		}
	}
	return float64(count), nil
}

func evalCOUNTIFS(args []interface{}) (interface{}, error) {
	if len(args) < 2 || len(args)%2 != 0 {
		return nil, fmt.Errorf("COUNTIFS requires pairs of (range, criteria)")
	}
	pairs := len(args) / 2
	ranges := make([][]interface{}, pairs)
	criterias := make([]string, pairs)
	for i := 0; i < pairs; i++ {
		ranges[i] = flattenArg(args[i*2])
		if ranges[i] == nil {
			return float64(0), nil
		}
		criterias[i] = fmt.Sprintf("%v", args[i*2+1])
	}
	count := 0
	for idx := range ranges[0] {
		allMatch := true
		for p := 0; p < pairs; p++ {
			if idx >= len(ranges[p]) || !matchesCriteria(ranges[p][idx], criterias[p]) {
				allMatch = false
				break
			}
		}
		if allMatch {
			count++
		}
	}
	return float64(count), nil
}

func evalSUMIFS(args []interface{}) (interface{}, error) {
	if len(args) < 3 || (len(args)-1)%2 != 0 {
		return nil, fmt.Errorf("SUMIFS requires sum_range + pairs of (range, criteria)")
	}
	sumVals := flattenArg(args[0])
	if sumVals == nil {
		return float64(0), nil
	}
	pairs := (len(args) - 1) / 2
	ranges := make([][]interface{}, pairs)
	criterias := make([]string, pairs)
	for i := 0; i < pairs; i++ {
		ranges[i] = flattenArg(args[1+i*2])
		if ranges[i] == nil {
			return float64(0), nil
		}
		criterias[i] = fmt.Sprintf("%v", args[2+i*2])
	}
	sum := 0.0
	for idx := range ranges[0] {
		allMatch := true
		for p := 0; p < pairs; p++ {
			if idx >= len(ranges[p]) || !matchesCriteria(ranges[p][idx], criterias[p]) {
				allMatch = false
				break
			}
		}
		if allMatch && idx < len(sumVals) {
			if f, ok := toFloat(sumVals[idx]); ok {
				sum += f
			}
		}
	}
	return sum, nil
}

func valuesMatch(a, b interface{}) bool {
	if fmt.Sprintf("%v", a) == fmt.Sprintf("%v", b) {
		return true
	}
	af, aok := toFloat(a)
	bf, bok := toFloat(b)
	return aok && bok && af == bf
}

func evalVLOOKUP(args []interface{}) (interface{}, error) {
	if len(args) < 3 {
		return nil, fmt.Errorf("VLOOKUP requires at least 3 arguments")
	}
	searchVal := args[0]
	rd, ok := args[1].(*rangeData)
	if !ok {
		return nil, fmt.Errorf("VLOOKUP: second argument must be a range")
	}
	colIdx, cok := toFloat(args[2])
	if !cok {
		return nil, fmt.Errorf("VLOOKUP: third argument must be numeric")
	}
	colIndex := int(colIdx)
	if colIndex < 1 || colIndex > rd.cols {
		return nil, fmt.Errorf("VLOOKUP: column index out of range")
	}
	exactMatch := false
	if len(args) >= 4 {
		exactMatch = !toBool(args[3])
	}
	_ = exactMatch
	for ri := 0; ri < rd.rows; ri++ {
		cellVal := rd.values[ri*rd.cols]
		if valuesMatch(searchVal, cellVal) {
			return rd.values[ri*rd.cols+colIndex-1], nil
		}
	}
	return nil, fmt.Errorf("VLOOKUP: value not found")
}

func evalHLOOKUP(args []interface{}) (interface{}, error) {
	if len(args) < 3 {
		return nil, fmt.Errorf("HLOOKUP requires at least 3 arguments")
	}
	searchVal := args[0]
	rd, ok := args[1].(*rangeData)
	if !ok {
		return nil, fmt.Errorf("HLOOKUP: second argument must be a range")
	}
	rowIdx, rok := toFloat(args[2])
	if !rok {
		return nil, fmt.Errorf("HLOOKUP: third argument must be numeric")
	}
	rowIndex := int(rowIdx)
	if rowIndex < 1 || rowIndex > rd.rows {
		return nil, fmt.Errorf("HLOOKUP: row index out of range")
	}
	exactMatch := false
	if len(args) >= 4 {
		exactMatch = !toBool(args[3])
	}
	_ = exactMatch
	for ci := 0; ci < rd.cols; ci++ {
		cellVal := rd.values[ci]
		if valuesMatch(searchVal, cellVal) {
			return rd.values[(rowIndex-1)*rd.cols+ci], nil
		}
	}
	return nil, fmt.Errorf("HLOOKUP: value not found")
}

func evalINDEX(args []interface{}) (interface{}, error) {
	if len(args) < 2 {
		return nil, fmt.Errorf("INDEX requires at least 2 arguments")
	}
	rd, ok := args[0].(*rangeData)
	if !ok {
		return nil, fmt.Errorf("INDEX: first argument must be a range")
	}
	rowNum, rok := toFloat(args[1])
	if !rok {
		return nil, fmt.Errorf("INDEX: row number must be numeric")
	}
	colNum := 1.0
	if len(args) >= 3 {
		cn, cnok := toFloat(args[2])
		if !cnok {
			return nil, fmt.Errorf("INDEX: column number must be numeric")
		}
		colNum = cn
	}
	ri := int(rowNum) - 1
	ci := int(colNum) - 1
	if ri < 0 || ri >= rd.rows || ci < 0 || ci >= rd.cols {
		return nil, fmt.Errorf("INDEX: out of range")
	}
	return rd.values[ri*rd.cols+ci], nil
}

func evalMATCH(args []interface{}) (interface{}, error) {
	if len(args) < 2 {
		return nil, fmt.Errorf("MATCH requires at least 2 arguments")
	}
	searchVal := args[0]
	vals := flattenArg(args[1])
	if vals == nil {
		return nil, fmt.Errorf("MATCH: second argument must be a range")
	}
	for i, v := range vals {
		if fmt.Sprintf("%v", v) == fmt.Sprintf("%v", searchVal) {
			return float64(i + 1), nil
		}
		sf, sok := toFloat(searchVal)
		cf, cfok := toFloat(v)
		if sok && cfok && sf == cf {
			return float64(i + 1), nil
		}
	}
	return nil, fmt.Errorf("MATCH: value not found")
}

func evalDATE(args []interface{}) (interface{}, error) {
	if len(args) != 3 {
		return nil, fmt.Errorf("DATE requires 3 arguments")
	}
	y, ok1 := toFloat(args[0])
	m, ok2 := toFloat(args[1])
	d, ok3 := toFloat(args[2])
	if !ok1 || !ok2 || !ok3 {
		return nil, fmt.Errorf("DATE requires numeric arguments")
	}
	t := time.Date(int(y), time.Month(int(m)), int(d), 0, 0, 0, 0, time.UTC)
	epoch := time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)
	days := t.Sub(epoch).Hours() / 24
	return days, nil
}

func evalTODAY(args []interface{}) (interface{}, error) {
	now := time.Now()
	t := time.Date(now.Year(), now.Month(), now.Day(), 0, 0, 0, 0, time.UTC)
	epoch := time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)
	return t.Sub(epoch).Hours() / 24, nil
}

func evalNOW(args []interface{}) (interface{}, error) {
	now := time.Now().UTC()
	epoch := time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)
	return now.Sub(epoch).Hours() / 24, nil
}

func evalYEAR(args []interface{}) (interface{}, error) {
	if len(args) != 1 {
		return nil, fmt.Errorf("YEAR requires 1 argument")
	}
	f, ok := toFloat(args[0])
	if !ok {
		return nil, fmt.Errorf("YEAR requires numeric argument")
	}
	epoch := time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)
	t := epoch.Add(time.Duration(f*24) * time.Hour)
	return float64(t.Year()), nil
}

func evalMONTH(args []interface{}) (interface{}, error) {
	if len(args) != 1 {
		return nil, fmt.Errorf("MONTH requires 1 argument")
	}
	f, ok := toFloat(args[0])
	if !ok {
		return nil, fmt.Errorf("MONTH requires numeric argument")
	}
	epoch := time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)
	t := epoch.Add(time.Duration(f*24) * time.Hour)
	return float64(t.Month()), nil
}

func evalDAY(args []interface{}) (interface{}, error) {
	if len(args) != 1 {
		return nil, fmt.Errorf("DAY requires 1 argument")
	}
	f, ok := toFloat(args[0])
	if !ok {
		return nil, fmt.Errorf("DAY requires numeric argument")
	}
	epoch := time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)
	t := epoch.Add(time.Duration(f*24) * time.Hour)
	return float64(t.Day()), nil
}

func evalFIND(args []interface{}) (interface{}, error) {
	if len(args) < 2 {
		return nil, fmt.Errorf("FIND requires at least 2 arguments")
	}
	search := fmt.Sprintf("%v", args[0])
	text := fmt.Sprintf("%v", args[1])
	startPos := 0
	if len(args) >= 3 {
		sp, ok := toFloat(args[2])
		if ok {
			startPos = int(sp) - 1
		}
	}
	if startPos < 0 {
		startPos = 0
	}
	if startPos >= len(text) {
		return nil, fmt.Errorf("FIND: not found")
	}
	idx := strings.Index(text[startPos:], search)
	if idx < 0 {
		return nil, fmt.Errorf("FIND: not found")
	}
	return float64(startPos + idx + 1), nil
}

func evalSEARCH(args []interface{}) (interface{}, error) {
	if len(args) < 2 {
		return nil, fmt.Errorf("SEARCH requires at least 2 arguments")
	}
	search := strings.ToLower(fmt.Sprintf("%v", args[0]))
	text := strings.ToLower(fmt.Sprintf("%v", args[1]))
	startPos := 0
	if len(args) >= 3 {
		sp, ok := toFloat(args[2])
		if ok {
			startPos = int(sp) - 1
		}
	}
	if startPos < 0 {
		startPos = 0
	}
	if startPos >= len(text) {
		return nil, fmt.Errorf("SEARCH: not found")
	}
	idx := strings.Index(text[startPos:], search)
	if idx < 0 {
		return nil, fmt.Errorf("SEARCH: not found")
	}
	return float64(startPos + idx + 1), nil
}

func evalSUBSTITUTE(args []interface{}) (interface{}, error) {
	if len(args) < 3 {
		return nil, fmt.Errorf("SUBSTITUTE requires at least 3 arguments")
	}
	text := fmt.Sprintf("%v", args[0])
	old := fmt.Sprintf("%v", args[1])
	newStr := fmt.Sprintf("%v", args[2])
	if len(args) >= 4 {
		inst, ok := toFloat(args[3])
		if ok && int(inst) > 0 {
			count := 0
			target := int(inst)
			result := strings.Builder{}
			remaining := text
			for {
				idx := strings.Index(remaining, old)
				if idx < 0 {
					result.WriteString(remaining)
					break
				}
				count++
				if count == target {
					result.WriteString(remaining[:idx])
					result.WriteString(newStr)
					result.WriteString(remaining[idx+len(old):])
					break
				}
				result.WriteString(remaining[:idx+len(old)])
				remaining = remaining[idx+len(old):]
			}
			return result.String(), nil
		}
	}
	return strings.ReplaceAll(text, old, newStr), nil
}

func evalREPLACE(args []interface{}) (interface{}, error) {
	if len(args) != 4 {
		return nil, fmt.Errorf("REPLACE requires 4 arguments")
	}
	text := fmt.Sprintf("%v", args[0])
	start, ok1 := toFloat(args[1])
	numChars, ok2 := toFloat(args[2])
	newText := fmt.Sprintf("%v", args[3])
	if !ok1 || !ok2 {
		return nil, fmt.Errorf("REPLACE requires numeric arguments for start and numChars")
	}
	s := int(start) - 1
	n := int(numChars)
	if s < 0 {
		s = 0
	}
	if s > len(text) {
		s = len(text)
	}
	end := s + n
	if end > len(text) {
		end = len(text)
	}
	return text[:s] + newText + text[end:], nil
}

func evalTEXT(args []interface{}) (interface{}, error) {
	if len(args) != 2 {
		return nil, fmt.Errorf("TEXT requires 2 arguments")
	}
	f, ok := toFloat(args[0])
	format := fmt.Sprintf("%v", args[1])
	if !ok {
		return fmt.Sprintf("%v", args[0]), nil
	}
	switch format {
	case "0":
		return strconv.FormatFloat(math.Round(f), 'f', 0, 64), nil
	case "0.00":
		return strconv.FormatFloat(f, 'f', 2, 64), nil
	case "0%":
		return strconv.FormatFloat(f*100, 'f', 0, 64) + "%", nil
	case "0.00%":
		return strconv.FormatFloat(f*100, 'f', 2, 64) + "%", nil
	default:
		return strconv.FormatFloat(f, 'f', -1, 64), nil
	}
}

func evalVALUE(args []interface{}) (interface{}, error) {
	if len(args) != 1 {
		return nil, fmt.Errorf("VALUE requires 1 argument")
	}
	s := fmt.Sprintf("%v", args[0])
	f, err := strconv.ParseFloat(strings.TrimSpace(s), 64)
	if err != nil {
		return nil, fmt.Errorf("VALUE: cannot convert %q to number", s)
	}
	return f, nil
}

func evalSUMPRODUCT(args []interface{}) (interface{}, error) {
	if len(args) == 0 {
		return float64(0), nil
	}
	firstVals := flattenArg(args[0])
	if firstVals == nil {
		return float64(0), nil
	}
	products := make([]float64, len(firstVals))
	for i, v := range firstVals {
		f, _ := toFloat(v)
		products[i] = f
	}
	for ai := 1; ai < len(args); ai++ {
		vals := flattenArg(args[ai])
		if vals == nil {
			continue
		}
		for i := 0; i < len(products) && i < len(vals); i++ {
			f, _ := toFloat(vals[i])
			products[i] *= f
		}
	}
	sum := 0.0
	for _, p := range products {
		sum += p
	}
	return sum, nil
}

func evalComparison(left interface{}, op string, right interface{}) bool {
	lf, lok := toFloat(left)
	rf, rok := toFloat(right)

	if lok && rok {
		switch op {
		case "=":
			return lf == rf
		case "<>":
			return lf != rf
		case "<":
			return lf < rf
		case ">":
			return lf > rf
		case "<=":
			return lf <= rf
		case ">=":
			return lf >= rf
		}
	}

	ls := fmt.Sprintf("%v", left)
	rs := fmt.Sprintf("%v", right)

	switch op {
	case "=":
		return ls == rs
	case "<>":
		return ls != rs
	case "<":
		return ls < rs
	case ">":
		return ls > rs
	case "<=":
		return ls <= rs
	case ">=":
		return ls >= rs
	}

	return false
}

func toFloat(v interface{}) (float64, bool) {
	switch val := v.(type) {
	case float64:
		return val, true
	case int:
		return float64(val), true
	case int64:
		return float64(val), true
	case string:
		if val == "" {
			return 0, true
		}
		if f, err := strconv.ParseFloat(val, 64); err == nil {
			return f, true
		}
		return 0, false
	case bool:
		if val {
			return 1, true
		}
		return 0, true
	}
	return 0, false
}

func toBool(v interface{}) bool {
	switch val := v.(type) {
	case bool:
		return val
	case float64:
		return val != 0
	case int:
		return val != 0
	case string:
		return val != "" && val != "0" && strings.ToUpper(val) != "FALSE"
	}
	return false
}

func extractCellName(ref string) string {
	result := strings.Builder{}
	for _, c := range ref {
		if unicode.IsLetter(c) || unicode.IsDigit(c) {
			result.WriteRune(unicode.ToUpper(c))
		}
	}
	return result.String()
}

func splitRef(ref string) (col, rowStr string) {
	for _, c := range ref {
		if unicode.IsLetter(c) {
			col += string(unicode.ToUpper(c))
		} else if unicode.IsDigit(c) {
			rowStr += string(c)
		}
	}
	return
}
