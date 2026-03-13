package goods

import (
	"fmt"
	"math"
	"strconv"
	"strings"
	"unicode"
)

type CellValues map[string]interface{}

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

	return Evaluate(c.formula, values)
}

type formulaParser struct {
	input  string
	pos    int
	values CellValues
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
		lf, lok := toFloat(left)
		rf, rok := toFloat(right)
		if lok && rok {
			if ch == '+' {
				left = lf + rf
			} else {
				left = lf - rf
			}
		} else if ch == '+' {
			left = fmt.Sprintf("%v%v", left, right)
		} else {
			return nil, fmt.Errorf("cannot subtract non-numeric values")
		}
	}

	return left, nil
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

func (p *formulaParser) parseCellReference() (interface{}, error) {
	p.pos++
	if p.pos < len(p.input) && p.input[p.pos] == '.' {
		p.pos++
	}

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
		return p.evalRange(ref, endRef)
	}

	if p.pos < len(p.input) {
		p.pos++
	}

	cellName := extractCellName(ref)
	val, ok := p.values[cellName]
	if !ok {
		return float64(0), nil
	}
	return val, nil
}

func (p *formulaParser) evalRange(startRef, endRef string) ([]interface{}, error) {
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

	var vals []interface{}

	if startRowNum > 0 && endRowNum > 0 && startRowNum != endRowNum {
		for ri := startRowNum; ri <= endRowNum; ri++ {
			for ci := startColIdx; ci <= endColIdx; ci++ {
				name := columnNumberToName(ci) + strconv.Itoa(ri)
				if v, ok := p.values[name]; ok {
					vals = append(vals, v)
				} else {
					vals = append(vals, float64(0))
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
				vals = append(vals, float64(0))
			}
		}
	}

	return vals, nil
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

	return evalFunction(name, args)
}

func evalFunction(name string, args []interface{}) (interface{}, error) {
	switch name {
	case "IF":
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

	case "AND":
		for _, a := range args {
			if !toBool(a) {
				return false, nil
			}
		}
		return true, nil

	case "OR":
		for _, a := range args {
			if toBool(a) {
				return true, nil
			}
		}
		return false, nil

	case "NOT":
		if len(args) != 1 {
			return nil, fmt.Errorf("NOT requires 1 argument")
		}
		return !toBool(args[0]), nil

	case "SUM":
		sum := 0.0
		for _, a := range args {
			if vals, ok := a.([]interface{}); ok {
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

	case "AVERAGE":
		sum := 0.0
		count := 0
		for _, a := range args {
			if vals, ok := a.([]interface{}); ok {
				for _, v := range vals {
					if f, ok := toFloat(v); ok {
						sum += f
						count++
					}
				}
			} else if f, ok := toFloat(a); ok {
				sum += f
				count++
			}
		}
		if count == 0 {
			return nil, fmt.Errorf("AVERAGE: no numeric values")
		}
		return sum / float64(count), nil

	case "COUNT":
		count := 0
		for _, a := range args {
			if vals, ok := a.([]interface{}); ok {
				for _, v := range vals {
					if _, ok := toFloat(v); ok {
						count++
					}
				}
			} else if _, ok := toFloat(a); ok {
				count++
			}
		}
		return float64(count), nil

	case "ABS":
		if len(args) != 1 {
			return nil, fmt.Errorf("ABS requires 1 argument")
		}
		if f, ok := toFloat(args[0]); ok {
			return math.Abs(f), nil
		}
		return nil, fmt.Errorf("ABS requires numeric argument")

	case "MIN":
		minVal := math.Inf(1)
		found := false
		for _, a := range args {
			if vals, ok := a.([]interface{}); ok {
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

	case "MAX":
		maxVal := math.Inf(-1)
		found := false
		for _, a := range args {
			if vals, ok := a.([]interface{}); ok {
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

	case "ROUND":
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

	case "FLOOR":
		if len(args) < 1 {
			return nil, fmt.Errorf("FLOOR requires at least 1 argument")
		}
		f, ok := toFloat(args[0])
		if !ok {
			return nil, fmt.Errorf("FLOOR requires numeric argument")
		}
		return math.Floor(f), nil

	case "CEIL", "CEILING":
		if len(args) < 1 {
			return nil, fmt.Errorf("CEILING requires at least 1 argument")
		}
		f, ok := toFloat(args[0])
		if !ok {
			return nil, fmt.Errorf("CEILING requires numeric argument")
		}
		return math.Ceil(f), nil

	case "CONCATENATE":
		var sb strings.Builder
		for _, a := range args {
			sb.WriteString(fmt.Sprintf("%v", a))
		}
		return sb.String(), nil

	case "LEN":
		if len(args) != 1 {
			return nil, fmt.Errorf("LEN requires 1 argument")
		}
		return float64(len(fmt.Sprintf("%v", args[0]))), nil

	case "TRIM":
		if len(args) != 1 {
			return nil, fmt.Errorf("TRIM requires 1 argument")
		}
		return strings.TrimSpace(fmt.Sprintf("%v", args[0])), nil

	case "UPPER":
		if len(args) != 1 {
			return nil, fmt.Errorf("UPPER requires 1 argument")
		}
		return strings.ToUpper(fmt.Sprintf("%v", args[0])), nil

	case "LOWER":
		if len(args) != 1 {
			return nil, fmt.Errorf("LOWER requires 1 argument")
		}
		return strings.ToLower(fmt.Sprintf("%v", args[0])), nil

	case "LEFT":
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

	case "RIGHT":
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

	case "MID":
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

	case "MOD":
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

	case "POWER":
		if len(args) != 2 {
			return nil, fmt.Errorf("POWER requires 2 arguments")
		}
		base, ok1 := toFloat(args[0])
		exp, ok2 := toFloat(args[1])
		if !ok1 || !ok2 {
			return nil, fmt.Errorf("POWER requires numeric arguments")
		}
		return math.Pow(base, exp), nil

	case "SQRT":
		if len(args) != 1 {
			return nil, fmt.Errorf("SQRT requires 1 argument")
		}
		f, ok := toFloat(args[0])
		if !ok {
			return nil, fmt.Errorf("SQRT requires numeric argument")
		}
		return math.Sqrt(f), nil
	}

	return nil, fmt.Errorf("unknown function: %s", name)
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
