package goods

import "errors"

var (
	ErrSheetNotFound     = errors.New("goods: sheet not found")
	ErrSheetExists       = errors.New("goods: sheet already exists")
	ErrSheetNameEmpty    = errors.New("goods: sheet name cannot be empty")
	ErrNoSheets          = errors.New("goods: workbook must have at least one sheet")
	ErrInvalidCell       = errors.New("goods: invalid cell reference")
	ErrInvalidCoords     = errors.New("goods: invalid coordinates")
	ErrColumnOutOfRange  = errors.New("goods: column number out of range")
	ErrRowOutOfRange     = errors.New("goods: row number out of range")
	ErrMergeOverlap      = errors.New("goods: merge range overlaps with existing merge")
	ErrMergeNotFound     = errors.New("goods: merge range not found")
	ErrStyleNotFound     = errors.New("goods: style not found")
	ErrFileClosed        = errors.New("goods: file is closed")
	ErrInvalidODS        = errors.New("goods: invalid ODS file")
	ErrUnsupportedType   = errors.New("goods: unsupported value type")
	ErrCircularReference = errors.New("goods: circular reference detected")
)
