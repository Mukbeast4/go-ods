package goods

import (
	"encoding/csv"
	"io"
	"os"
)

type CSVOptions struct {
	Separator rune
	UseCRLF   bool
}

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

func (f *File) ExportCSVFile(sheet, path string, opts *CSVOptions) error {
	file, err := os.Create(path)
	if err != nil {
		return err
	}
	defer file.Close()

	return f.ExportCSV(sheet, file, opts)
}
