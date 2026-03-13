package goods

import (
	"fmt"
	"strings"
)

type dataValidation struct {
	name       string
	validation *DataValidation
}

type DataValidation struct {
	Type         string
	Operator     string
	Formula1     string
	Formula2     string
	AllowEmpty   bool
	ErrorTitle   string
	ErrorMessage string
	ErrorStyle   string
	InputTitle   string
	InputMessage string
}

func (f *File) SetDataValidation(sheet, topLeft, bottomRight string, dv *DataValidation) error {
	if f.closed {
		return ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
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

	name := fmt.Sprintf("val%d", len(s.validations)+1)

	dvCopy := *dv
	s.validations = append(s.validations, &dataValidation{
		name:       name,
		validation: &dvCopy,
	})

	for rowIdx := sr; rowIdx <= er; rowIdx++ {
		r := s.getOrCreateRow(rowIdx)
		for colIdx := sc; colIdx <= ec; colIdx++ {
			c, ok := r.cells[colIdx]
			if !ok {
				c = &cell{valueType: CellTypeEmpty}
				r.cells[colIdx] = c
			}
			c.validationName = name
		}
	}

	return nil
}

func (f *File) GetDataValidation(sheet, cellRef string) (*DataValidation, error) {
	if f.closed {
		return nil, ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
		return nil, ErrSheetNotFound
	}

	col, row, err := CellNameToCoordinates(cellRef)
	if err != nil {
		return nil, err
	}

	c := s.getCell(col, row)
	if c == nil || c.validationName == "" {
		return nil, nil
	}

	for _, dv := range s.validations {
		if dv.name == c.validationName {
			result := *dv.validation
			return &result, nil
		}
	}

	return nil, nil
}

func (f *File) RemoveDataValidation(sheet, topLeft, bottomRight string) error {
	if f.closed {
		return ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
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

	removedNames := make(map[string]bool)

	for rowIdx := sr; rowIdx <= er; rowIdx++ {
		r, ok := s.rows[rowIdx]
		if !ok {
			continue
		}
		for colIdx := sc; colIdx <= ec; colIdx++ {
			c, ok := r.cells[colIdx]
			if !ok {
				continue
			}
			if c.validationName != "" {
				removedNames[c.validationName] = true
				c.validationName = ""
			}
		}
	}

	newValidations := make([]*dataValidation, 0, len(s.validations))
	for _, dv := range s.validations {
		if !removedNames[dv.name] {
			newValidations = append(newValidations, dv)
		}
	}
	s.validations = newValidations

	return nil
}

func buildValidationCondition(dv *DataValidation) string {
	switch dv.Type {
	case "list":
		return fmt.Sprintf("of:cell-content-is-in-list(%s)", dv.Formula1)
	case "whole-number", "decimal", "text-length":
		return buildNumberCondition(dv)
	case "date", "time":
		return buildNumberCondition(dv)
	case "custom":
		return fmt.Sprintf("of:is-true-formula(%s)", dv.Formula1)
	default:
		return ""
	}
}

func buildNumberCondition(dv *DataValidation) string {
	switch dv.Operator {
	case "between":
		return fmt.Sprintf("of:cell-content-is-between(%s,%s)", dv.Formula1, dv.Formula2)
	case "not-between":
		return fmt.Sprintf("of:cell-content-is-not-between(%s,%s)", dv.Formula1, dv.Formula2)
	case "equal":
		return fmt.Sprintf("of:cell-content()=%s", dv.Formula1)
	case "not-equal":
		return fmt.Sprintf("of:cell-content()!=%s", dv.Formula1)
	case "greater-than":
		return fmt.Sprintf("of:cell-content()>%s", dv.Formula1)
	case "less-than":
		return fmt.Sprintf("of:cell-content()<%s", dv.Formula1)
	case "greater-than-or-equal":
		return fmt.Sprintf("of:cell-content()>=%s", dv.Formula1)
	case "less-than-or-equal":
		return fmt.Sprintf("of:cell-content()<=%s", dv.Formula1)
	default:
		return ""
	}
}

func parseValidationCondition(condition string) (dvType, operator, formula1, formula2 string) {
	condition = strings.TrimSpace(condition)

	if strings.HasPrefix(condition, "of:cell-content-is-in-list(") {
		inner := condition[len("of:cell-content-is-in-list(") : len(condition)-1]
		return "list", "", inner, ""
	}

	if strings.HasPrefix(condition, "of:is-true-formula(") {
		inner := condition[len("of:is-true-formula(") : len(condition)-1]
		return "custom", "", inner, ""
	}

	if strings.HasPrefix(condition, "of:cell-content-is-between(") {
		inner := condition[len("of:cell-content-is-between(") : len(condition)-1]
		parts := strings.SplitN(inner, ",", 2)
		if len(parts) == 2 {
			return "whole-number", "between", strings.TrimSpace(parts[0]), strings.TrimSpace(parts[1])
		}
	}

	if strings.HasPrefix(condition, "of:cell-content-is-not-between(") {
		inner := condition[len("of:cell-content-is-not-between(") : len(condition)-1]
		parts := strings.SplitN(inner, ",", 2)
		if len(parts) == 2 {
			return "whole-number", "not-between", strings.TrimSpace(parts[0]), strings.TrimSpace(parts[1])
		}
	}

	if strings.HasPrefix(condition, "of:cell-content()") {
		rest := condition[len("of:cell-content()"):]
		for _, op := range []struct {
			prefix string
			name   string
		}{
			{">=", "greater-than-or-equal"},
			{"<=", "less-than-or-equal"},
			{"!=", "not-equal"},
			{">", "greater-than"},
			{"<", "less-than"},
			{"=", "equal"},
		} {
			if strings.HasPrefix(rest, op.prefix) {
				return "whole-number", op.name, strings.TrimSpace(rest[len(op.prefix):]), ""
			}
		}
	}

	return "", "", condition, ""
}
