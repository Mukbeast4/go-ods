package goods

import "errors"

var (
	ErrSheetNotFound      = errors.New("goods: sheet not found")
	ErrSheetExists        = errors.New("goods: sheet already exists")
	ErrSheetNameEmpty     = errors.New("goods: sheet name cannot be empty")
	ErrNoSheets           = errors.New("goods: workbook must have at least one sheet")
	ErrInvalidCell        = errors.New("goods: invalid cell reference")
	ErrInvalidCoords      = errors.New("goods: invalid coordinates")
	ErrColumnOutOfRange   = errors.New("goods: column number out of range")
	ErrRowOutOfRange      = errors.New("goods: row number out of range")
	ErrMergeOverlap       = errors.New("goods: merge range overlaps with existing merge")
	ErrMergeNotFound      = errors.New("goods: merge range not found")
	ErrStyleNotFound      = errors.New("goods: style not found")
	ErrFileClosed         = errors.New("goods: file is closed")
	ErrInvalidODS         = errors.New("goods: invalid ODS file")
	ErrUnsupportedType    = errors.New("goods: unsupported value type")
	ErrCircularReference  = errors.New("goods: circular reference detected")
	ErrNamedRangeNotFound = errors.New("goods: named range not found")
	ErrNamedRangeExists   = errors.New("goods: named range already exists")
	ErrValidationNotFound = errors.New("goods: data validation not found")
	ErrAutoFilterExists   = errors.New("goods: auto-filter already exists on this sheet")
	ErrAutoFilterNotFound = errors.New("goods: auto-filter not found")
)
