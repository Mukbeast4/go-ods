package goods

import (
	"bytes"
	"fmt"
	"io"
	"os"

	oxml "github.com/mukbeast4/go-ods/internal/xml"
	ozip "github.com/mukbeast4/go-ods/internal/zip"
)

// Save writes the document back to the path it was opened from. It returns an
// error when the file was built with NewFile and has no path yet; use SaveAs
// instead.
func (f *File) Save() error {
	if f.closed {
		return ErrFileClosed
	}
	if f.path == "" {
		return fmt.Errorf("goods: no file path set, use SaveAs")
	}
	return f.SaveAs(f.path)
}

// SaveAs writes the document to path and remembers it as the current path for
// subsequent Save calls.
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

// Write serializes the document and writes the resulting zip archive to w.
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

// WriteToBuffer serializes the document into a new bytes.Buffer.
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

// Close releases the in-memory data. Once closed a File cannot be read or
// written; every method returns ErrFileClosed.
func (f *File) Close() error {
	f.closed = true
	f.sheets = nil
	f.rawFiles = nil
	f.styles = nil
	return nil
}

// SaveAndClose is a convenience that calls SaveAs followed by Close.
func (f *File) SaveAndClose(path string) error {
	if err := f.SaveAs(path); err != nil {
		return err
	}
	return f.Close()
}

// WriteToFile creates path and writes the serialized document to it.
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

	settingsData, err := f.marshalSettings()
	if err != nil {
		return nil, fmt.Errorf("marshal settings: %w", err)
	}
	if settingsData != nil {
		entries = append(entries, ozip.WriteEntry{Name: "settings.xml", Data: settingsData})
	}

	for name, data := range f.rawFiles {
		entries = append(entries, ozip.WriteEntry{Name: name, Data: data})
	}

	for name, data := range f.images {
		entries = append(entries, ozip.WriteEntry{Name: name, Data: data})
	}

	return entries, nil
}
