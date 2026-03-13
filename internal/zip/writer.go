package zip

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
	"os"
)

type WriteEntry struct {
	Name string
	Data []byte
}

func WriteFile(path string, entries []WriteEntry) error {
	f, err := os.Create(path)
	if err != nil {
		return fmt.Errorf("create file: %w", err)
	}
	defer f.Close()

	return WriteTo(f, entries)
}

func WriteToBuffer(entries []WriteEntry) (*bytes.Buffer, error) {
	buf := new(bytes.Buffer)
	if err := WriteTo(buf, entries); err != nil {
		return nil, err
	}
	return buf, nil
}

func WriteTo(w io.Writer, entries []WriteEntry) error {
	zw := zip.NewWriter(w)
	defer zw.Close()

	for i, entry := range entries {
		header := &zip.FileHeader{
			Name:   entry.Name,
			Method: zip.Deflate,
		}

		if i == 0 && entry.Name == "mimetype" {
			header.Method = zip.Store
		}

		writer, err := zw.CreateHeader(header)
		if err != nil {
			return fmt.Errorf("create entry %s: %w", entry.Name, err)
		}

		if _, err := writer.Write(entry.Data); err != nil {
			return fmt.Errorf("write entry %s: %w", entry.Name, err)
		}
	}

	return nil
}
