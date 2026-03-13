package zip

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
	"os"
)

type ReadResult struct {
	Files map[string][]byte
}

func ReadFile(path string) (*ReadResult, error) {
	data, err := os.ReadFile(path)
	if err != nil {
		return nil, fmt.Errorf("read file: %w", err)
	}
	return ReadBytes(data)
}

func ReadBytes(data []byte) (*ReadResult, error) {
	r, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return nil, fmt.Errorf("open zip: %w", err)
	}
	return readFromZipReader(r)
}

func ReadFromReader(r io.ReaderAt, size int64) (*ReadResult, error) {
	zr, err := zip.NewReader(r, size)
	if err != nil {
		return nil, fmt.Errorf("open zip: %w", err)
	}
	return readFromZipReader(zr)
}

func readFromZipReader(r *zip.Reader) (*ReadResult, error) {
	result := &ReadResult{
		Files: make(map[string][]byte),
	}

	for _, f := range r.File {
		rc, err := f.Open()
		if err != nil {
			return nil, fmt.Errorf("open %s: %w", f.Name, err)
		}
		data, err := io.ReadAll(rc)
		rc.Close()
		if err != nil {
			return nil, fmt.Errorf("read %s: %w", f.Name, err)
		}
		result.Files[f.Name] = data
	}

	mimeData, ok := result.Files["mimetype"]
	if !ok {
		return nil, fmt.Errorf("missing mimetype file")
	}
	if string(bytes.TrimSpace(mimeData)) != "application/vnd.oasis.opendocument.spreadsheet" {
		return nil, fmt.Errorf("invalid mimetype: %s", string(mimeData))
	}

	if _, ok := result.Files["content.xml"]; !ok {
		return nil, fmt.Errorf("missing content.xml")
	}

	return result, nil
}
