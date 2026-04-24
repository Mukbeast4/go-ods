package goods

import (
	"encoding/csv"
	"io"
	"os"
)

// CSVOptions controls the encoding used by ExportCSV. Separator defaults to
// ',' when zero; set UseCRLF to true for Windows line endings.
type CSVOptions struct {
	Separator rune
	UseCRLF   bool
}

// ExportCSV writes the sheet contents as CSV to w. Formulas are exported as
// their computed values (or raw values when no recalc has been performed).
func (f *File) ExportCSV(sheet string, w io.Writer, opts *CSVOptions) error {
	if f.closed {
		return ErrFileClosed
	}

	rows, err := f.GetRows(sheet)
	if err != nil {
		return err
	}

	writer := csv.NewWriter(w)
	if opts != nil {
		if opts.Separator != 0 {
			writer.Comma = opts.Separator
		}
		writer.UseCRLF = opts.UseCRLF
	}

	for _, row := range rows {
		if err := writer.Write(row); err != nil {
			return err
		}
	}

	writer.Flush()
	return writer.Error()
}

// ExportCSVFile creates path and writes the sheet contents as CSV to it.
func (f *File) ExportCSVFile(sheet, path string, opts *CSVOptions) error {
	file, err := os.Create(path)
	if err != nil {
		return err
	}
	defer file.Close()

	return f.ExportCSV(sheet, file, opts)
}
