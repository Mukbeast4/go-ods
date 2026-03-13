package goods

import (
	"bytes"
	"fmt"
	"io"
	"os"

	oxml "github.com/mukbeast4/go-ods/internal/xml"
	ozip "github.com/mukbeast4/go-ods/internal/zip"
)

func (f *File) Save() error {
	if f.closed {
		return ErrFileClosed
	}
	if f.path == "" {
		return fmt.Errorf("goods: no file path set, use SaveAs")
	}
	return f.SaveAs(f.path)
}

func (f *File) SaveAs(path string) error {
	if f.closed {
		return ErrFileClosed
	}

	entries, err := f.buildZipEntries()
	if err != nil {
		return err
	}

	if err := ozip.WriteFile(path, entries); err != nil {
		return fmt.Errorf("save file: %w", err)
	}

	f.path = path
	return nil
}

func (f *File) Write(w io.Writer) error {
	if f.closed {
		return ErrFileClosed
	}

	entries, err := f.buildZipEntries()
	if err != nil {
		return err
	}

	return ozip.WriteTo(w, entries)
}

func (f *File) WriteToBuffer() (*bytes.Buffer, error) {
	if f.closed {
		return nil, ErrFileClosed
	}

	entries, err := f.buildZipEntries()
	if err != nil {
		return nil, err
	}

	return ozip.WriteToBuffer(entries)
}

func (f *File) Close() error {
	f.closed = true
	f.sheets = nil
	f.rawFiles = nil
	f.styles = nil
	return nil
}

func (f *File) SaveAndClose(path string) error {
	if err := f.SaveAs(path); err != nil {
		return err
	}
	return f.Close()
}

func (f *File) WriteToFile(path string) error {
	file, err := os.Create(path)
	if err != nil {
		return fmt.Errorf("create file: %w", err)
	}
	defer file.Close()

	return f.Write(file)
}

func (f *File) buildZipEntries() ([]ozip.WriteEntry, error) {
	contentData, err := f.marshalContent()
	if err != nil {
		return nil, fmt.Errorf("marshal content: %w", err)
	}

	metaData, err := f.marshalMeta()
	if err != nil {
		return nil, fmt.Errorf("marshal meta: %w", err)
	}

	stylesData, err := f.marshalStyles()
	if err != nil {
		return nil, fmt.Errorf("marshal styles: %w", err)
	}

	manifestData, err := f.marshalManifest()
	if err != nil {
		return nil, fmt.Errorf("marshal manifest: %w", err)
	}

	entries := []ozip.WriteEntry{
		{Name: "mimetype", Data: []byte(oxml.MimeTypeODS)},
		{Name: "content.xml", Data: contentData},
		{Name: "styles.xml", Data: stylesData},
		{Name: "meta.xml", Data: metaData},
		{Name: "META-INF/manifest.xml", Data: manifestData},
	}

	for name, data := range f.rawFiles {
		entries = append(entries, ozip.WriteEntry{Name: name, Data: data})
	}

	return entries, nil
}
