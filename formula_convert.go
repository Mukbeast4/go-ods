package goods

import (
	"fmt"
	"strings"
	"sync"
	"unicode"
)

func FormulaToGo(formula string) string {
	f := strings.TrimPrefix(formula, "of:=")
	f = strings.TrimPrefix(f, "of:")
	return convertExpr(f)
}

func convertExpr(expr string) string {
	expr = strings.TrimSpace(expr)

	if fn, args, ok := parseFuncExpr(expr); ok {
		return convertFuncToGo(fn, args)
	}

	expr = convertCellRefsToGo(expr)
	expr = convertOperatorsToGo(expr)

	return expr
}

func parseFuncExpr(expr string) (string, []string, bool) {
	parenIdx := strings.Index(expr, "(")
	if parenIdx < 0 {
		return "", nil, false
	}

	name := strings.TrimSpace(expr[:parenIdx])
	if name == "" || !isValidFuncNameGo(name) {
		return "", nil, false
	}

	if expr[len(expr)-1] != ')' {
		return "", nil, false
	}

	inner := expr[parenIdx+1 : len(expr)-1]
	args := splitFuncArgs(inner)

	return strings.ToUpper(name), args, true
}

func isValidFuncNameGo(name string) bool {
	for _, c := range name {
		if !unicode.IsLetter(c) && !unicode.IsDigit(c) && c != '_' {
			return false
		}
	}
	return true
}

func splitFuncArgs(s string) []string {
	var args []string
	depth := 0
	current := strings.Builder{}

	for _, c := range s {
		switch c {
		case '(':
			depth++
			current.WriteRune(c)
		case ')':
			depth--
			current.WriteRune(c)
		case ';':
			if depth == 0 {
				args = append(args, strings.TrimSpace(current.String()))
				current.Reset()
			} else {
				current.WriteRune(c)
			}
		default:
			current.WriteRune(c)
		}
	}

	if current.Len() > 0 {
		args = append(args, strings.TrimSpace(current.String()))
	}

	return args
}

type convertFunc func([]string) string

var (
	convertFunctionsOnce sync.Once
	convertFunctionsMap  map[string]convertFunc
)

func getConvertFunctions() map[string]convertFunc {
	convertFunctionsOnce.Do(func() {
		convertFunctionsMap = map[string]convertFunc{
			"IF":          convertIF,
			"AND":         convertAND,
			"OR":          convertOR,
			"NOT":         convertNOT,
			"ABS":         convertABS,
			"SUM":         convertSUM,
			"MIN":         convertMIN,
			"MAX":         convertMAX,
			"ROUND":       convertROUND,
			"CONCATENATE": convertCONCATENATE,
			"LEN":         convertLEN,
			"TRIM":        convertTRIM,
			"UPPER":       convertUPPER,
			"LOWER":       convertLOWER,
			"LEFT":        convertLEFT,
			"RIGHT":       convertRIGHT,
			"MID":         convertMID,
			"MOD":         convertMOD,
			"POWER":       convertPOWER,
			"SQRT":        convertSQRT,
			"AVERAGE":     convertAVERAGE,
		}
	})
	return convertFunctionsMap
}

func convertFuncToGo(name string, args []string) string {
	if fn, ok := getConvertFunctions()[name]; ok {
		if result := fn(args); result != "" {
			return result
		}
	}

	return convertFallback(name, args)
}

func convertFallback(name string, args []string) string {
	parts := make([]string, len(args))
	for i, a := range args {
		parts[i] = convertExpr(a)
	}
	return strings.ToLower(name) + "(" + strings.Join(parts, ", ") + ")"
}

func convertAllArgs(args []string) []string {
	parts := make([]string, len(args))
	for i, a := range args {
		parts[i] = convertExpr(a)
	}
	return parts
}

func convertIF(args []string) string {
	if len(args) == 3 {
		cond := convertExpr(args[0])
		valTrue := convertExpr(args[1])
		valFalse := convertExpr(args[2])
		return fmt.Sprintf("func() interface{} { if %s { return %s } else { return %s } }()", cond, valTrue, valFalse)
	}
	if len(args) == 2 {
		cond := convertExpr(args[0])
		valTrue := convertExpr(args[1])
		return fmt.Sprintf("func() interface{} { if %s { return %s } else { return nil } }()", cond, valTrue)
	}
	return ""
}

func convertAND(args []string) string {
	return "(" + strings.Join(convertAllArgs(args), " && ") + ")"
}

func convertOR(args []string) string {
	return "(" + strings.Join(convertAllArgs(args), " || ") + ")"
}

func convertNOT(args []string) string {
	if len(args) == 1 {
		return "!" + convertExpr(args[0])
	}
	return ""
}

func convertABS(args []string) string {
	if len(args) == 1 {
		return "math.Abs(" + convertExpr(args[0]) + ")"
	}
	return ""
}

func convertSUM(args []string) string {
	if len(args) == 1 {
		return convertSumRangeToGo(args[0])
	}
	return "(" + strings.Join(convertAllArgs(args), " + ") + ")"
}

func convertMIN(args []string) string {
	return "min(" + strings.Join(convertAllArgs(args), ", ") + ")"
}

func convertMAX(args []string) string {
	return "max(" + strings.Join(convertAllArgs(args), ", ") + ")"
}

func convertROUND(args []string) string {
	if len(args) >= 1 {
		val := convertExpr(args[0])
		if len(args) == 2 {
			return fmt.Sprintf("math.Round(%s, %s)", val, convertExpr(args[1]))
		}
		return fmt.Sprintf("math.Round(%s)", val)
	}
	return ""
}

func convertCONCATENATE(args []string) string {
	return strings.Join(convertAllArgs(args), " + ")
}

func convertLEN(args []string) string {
	if len(args) == 1 {
		return "len(" + convertExpr(args[0]) + ")"
	}
	return ""
}

func convertTRIM(args []string) string {
	if len(args) == 1 {
		return "strings.TrimSpace(" + convertExpr(args[0]) + ")"
	}
	return ""
}

func convertUPPER(args []string) string {
	if len(args) == 1 {
		return "strings.ToUpper(" + convertExpr(args[0]) + ")"
	}
	return ""
}

func convertLOWER(args []string) string {
	if len(args) == 1 {
		return "strings.ToLower(" + convertExpr(args[0]) + ")"
	}
	return ""
}

func convertLEFT(args []string) string {
	if len(args) == 2 {
		return fmt.Sprintf("%s[:%s]", convertExpr(args[0]), convertExpr(args[1]))
	}
	return ""
}

func convertRIGHT(args []string) string {
	if len(args) == 2 {
		val := convertExpr(args[0])
		n := convertExpr(args[1])
		return fmt.Sprintf("%s[len(%s)-%s:]", val, val, n)
	}
	return ""
}

func convertMID(args []string) string {
	if len(args) == 3 {
		val := convertExpr(args[0])
		start := convertExpr(args[1])
		length := convertExpr(args[2])
		return fmt.Sprintf("%s[%s-1:%s-1+%s]", val, start, start, length)
	}
	return ""
}

func convertMOD(args []string) string {
	if len(args) == 2 {
		return fmt.Sprintf("math.Mod(%s, %s)", convertExpr(args[0]), convertExpr(args[1]))
	}
	return ""
}

func convertPOWER(args []string) string {
	if len(args) == 2 {
		return fmt.Sprintf("math.Pow(%s, %s)", convertExpr(args[0]), convertExpr(args[1]))
	}
	return ""
}

func convertSQRT(args []string) string {
	if len(args) == 1 {
		return fmt.Sprintf("math.Sqrt(%s)", convertExpr(args[0]))
	}
	return ""
}

func convertAVERAGE(args []string) string {
	parts := convertAllArgs(args)
	return fmt.Sprintf("((%s) / %d)", strings.Join(parts, " + "), len(parts))
}

func convertSumRangeToGo(arg string) string {
	arg = strings.TrimSpace(arg)

	if strings.HasPrefix(arg, "[.") && strings.Contains(arg, ":.") {
		inner := strings.TrimPrefix(arg, "[.")
		inner = strings.TrimSuffix(inner, "]")
		parts := strings.SplitN(inner, ":.", 2)
		if len(parts) == 2 {
			startCol, _ := splitRef(parts[0])
			endCol, _ := splitRef(parts[1])

			startIdx := columnNameToNumber(startCol)
			endIdx := columnNameToNumber(endCol)

			var terms []string
			for i := startIdx; i <= endIdx; i++ {
				terms = append(terms, "r."+columnNumberToName(i))
			}
			return strings.Join(terms, " + ")
		}
	}

	return convertExpr(arg)
}

func convertCellRefsToGo(expr string) string {
	result := strings.Builder{}
	i := 0

	for i < len(expr) {
		if expr[i] == '[' && i+1 < len(expr) && expr[i+1] == '.' {
			end := strings.Index(expr[i:], "]")
			if end < 0 {
				result.WriteByte(expr[i])
				i++
				continue
			}

			ref := expr[i+2 : i+end]

			if strings.Contains(ref, ":.") {
				parts := strings.SplitN(ref, ":.", 2)
				startCol, _ := splitRef(parts[0])
				endCol, _ := splitRef(parts[1])

				startIdx := columnNameToNumber(startCol)
				endIdx := columnNameToNumber(endCol)

				var terms []string
				for j := startIdx; j <= endIdx; j++ {
					terms = append(terms, "r."+columnNumberToName(j))
				}
				result.WriteString(strings.Join(terms, " + "))
			} else {
				col, _ := splitRef(ref)
				result.WriteString("r." + col)
			}

			i += end + 1
		} else {
			result.WriteByte(expr[i])
			i++
		}
	}

	return result.String()
}

func convertOperatorsToGo(expr string) string {
	expr = strings.ReplaceAll(expr, "<>", "!=")

	result := strings.Builder{}
	i := 0
	for i < len(expr) {
		if expr[i] == '=' {
			if i > 0 && (expr[i-1] == '<' || expr[i-1] == '>' || expr[i-1] == '!' || expr[i-1] == '=') {
				result.WriteByte(expr[i])
				i++
				continue
			}
			if i+1 < len(expr) && expr[i+1] == '=' {
				result.WriteString("==")
				i += 2
				continue
			}
			result.WriteString("==")
			i++
		} else {
			result.WriteByte(expr[i])
			i++
		}
	}

	return result.String()
}
