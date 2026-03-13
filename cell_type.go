package goods

import (
	"fmt"
	"strconv"
	"time"
)

type CellType int

const (
	CellTypeEmpty CellType = iota
	CellTypeString
	CellTypeFloat
	CellTypeDate
	CellTypeBool
	CellTypeCurrency
	CellTypePercentage
	CellTypeTime
)

var cellTypeToODS = map[CellType]string{
	CellTypeEmpty:      "",
	CellTypeString:     "string",
	CellTypeFloat:      "float",
	CellTypeDate:       "date",
	CellTypeBool:       "boolean",
	CellTypeCurrency:   "currency",
	CellTypePercentage: "percentage",
	CellTypeTime:       "time",
}

var odsToCellType = map[string]CellType{
	"":           CellTypeEmpty,
	"string":     CellTypeString,
	"float":      CellTypeFloat,
	"date":       CellTypeDate,
	"boolean":    CellTypeBool,
	"currency":   CellTypeCurrency,
	"percentage": CellTypePercentage,
	"time":       CellTypeTime,
}

func (ct CellType) odsName() string {
	return cellTypeToODS[ct]
}

func cellTypeFromODS(name string) CellType {
	if ct, ok := odsToCellType[name]; ok {
		return ct
	}
	return CellTypeString
}

func detectCellType(value interface{}) (CellType, string) {
	switch v := value.(type) {
	case nil:
		return CellTypeEmpty, ""
	case string:
		return CellTypeString, v
	case int:
		return CellTypeFloat, strconv.Itoa(v)
	case int8:
		return CellTypeFloat, strconv.FormatInt(int64(v), 10)
	case int16:
		return CellTypeFloat, strconv.FormatInt(int64(v), 10)
	case int32:
		return CellTypeFloat, strconv.FormatInt(int64(v), 10)
	case int64:
		return CellTypeFloat, strconv.FormatInt(v, 10)
	case uint:
		return CellTypeFloat, strconv.FormatUint(uint64(v), 10)
	case uint8:
		return CellTypeFloat, strconv.FormatUint(uint64(v), 10)
	case uint16:
		return CellTypeFloat, strconv.FormatUint(uint64(v), 10)
	case uint32:
		return CellTypeFloat, strconv.FormatUint(uint64(v), 10)
	case uint64:
		return CellTypeFloat, strconv.FormatUint(v, 10)
	case float32:
		return CellTypeFloat, strconv.FormatFloat(float64(v), 'f', -1, 32)
	case float64:
		return CellTypeFloat, strconv.FormatFloat(v, 'f', -1, 64)
	case bool:
		if v {
			return CellTypeBool, "true"
		}
		return CellTypeBool, "false"
	case time.Time:
		return CellTypeDate, v.Format("2006-01-02T15:04:05")
	case fmt.Stringer:
		return CellTypeString, v.String()
	default:
		return CellTypeString, fmt.Sprintf("%v", v)
	}
}

func cellValueToString(valueType CellType, rawValue string) string {
	switch valueType {
	case CellTypeBool:
		if rawValue == "true" {
			return "TRUE"
		}
		return "FALSE"
	default:
		return rawValue
	}
}

func parseFloat(raw string) (float64, error) {
	return strconv.ParseFloat(raw, 64)
}

func parseInt(raw string) (int64, error) {
	f, err := strconv.ParseFloat(raw, 64)
	if err != nil {
		return 0, err
	}
	return int64(f), nil
}

func parseBool(raw string) (bool, error) {
	return strconv.ParseBool(raw)
}

func parseDate(raw string) (time.Time, error) {
	formats := []string{
		"2006-01-02T15:04:05",
		"2006-01-02",
	}
	for _, f := range formats {
		if t, err := time.Parse(f, raw); err == nil {
			return t, nil
		}
	}
	return time.Time{}, fmt.Errorf("goods: cannot parse date %q", raw)
}
