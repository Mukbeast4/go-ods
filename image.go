package goods

import (
	"bytes"
	"crypto/sha1"
	"encoding/hex"
	"fmt"
	"os"
)

type ImageOptions struct {
	Width   float64
	Height  float64
	OffsetX float64
	OffsetY float64
}

type ImageInfo struct {
	CellRef string
	Name    string
	Format  string
	Width   float64
	Height  float64
	OffsetX float64
	OffsetY float64
	Data    []byte
}

func (f *File) AddImage(sheet, cellRef, path string, opts *ImageOptions) error {
	data, err := os.ReadFile(path)
	if err != nil {
		return fmt.Errorf("read image: %w", err)
	}
	return f.AddImageFromBytes(sheet, cellRef, data, opts)
}

func (f *File) AddImageFromBytes(sheet, cellRef string, data []byte, opts *ImageOptions) error {
	if f.closed {
		return ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
		return ErrSheetNotFound
	}

	format, err := detectImageFormat(data)
	if err != nil {
		return err
	}

	col, row, err := CellNameToCoordinates(cellRef)
	if err != nil {
		return err
	}

	if opts == nil {
		opts = &ImageOptions{}
	}
	width, height := opts.Width, opts.Height
	if width <= 0 {
		width = 3.0
	}
	if height <= 0 {
		height = 3.0
	}

	sum := sha1.Sum(data)
	href := fmt.Sprintf("Pictures/%s.%s", hex.EncodeToString(sum[:]), format)

	if f.images == nil {
		f.images = make(map[string][]byte)
	}
	f.images[href] = data

	r := s.getOrCreateRow(row)
	c, ok := r.cells[col]
	if !ok {
		c = &cell{valueType: CellTypeEmpty}
		r.cells[col] = c
	}

	f.nextImageID++
	frame := &imageFrame{
		name:    fmt.Sprintf("Image %d", f.nextImageID),
		href:    href,
		width:   width,
		height:  height,
		offsetX: opts.OffsetX,
		offsetY: opts.OffsetY,
		format:  format,
		data:    data,
	}
	c.images = append(c.images, frame)

	if col > s.maxCol {
		s.maxCol = col
	}
	if row > s.maxRow {
		s.maxRow = row
	}

	return nil
}

func (f *File) GetImages(sheet string) ([]ImageInfo, error) {
	if f.closed {
		return nil, ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
		return nil, ErrSheetNotFound
	}

	var result []ImageInfo
	for rowIdx := 1; rowIdx <= s.maxRow; rowIdx++ {
		r, ok := s.rows[rowIdx]
		if !ok {
			continue
		}
		for colIdx := 1; colIdx <= s.maxCol; colIdx++ {
			c, ok := r.cells[colIdx]
			if !ok || len(c.images) == 0 {
				continue
			}
			ref, err := CoordinatesToCellName(colIdx, rowIdx)
			if err != nil {
				continue
			}
			for _, img := range c.images {
				data := img.data
				if data == nil {
					data = f.images[img.href]
				}
				result = append(result, ImageInfo{
					CellRef: ref,
					Name:    img.name,
					Format:  img.format,
					Width:   img.width,
					Height:  img.height,
					OffsetX: img.offsetX,
					OffsetY: img.offsetY,
					Data:    data,
				})
			}
		}
	}
	return result, nil
}

func (f *File) RemoveImages(sheet, cellRef string) error {
	if f.closed {
		return ErrFileClosed
	}
	s := f.getSheet(sheet)
	if s == nil {
		return ErrSheetNotFound
	}

	col, row, err := CellNameToCoordinates(cellRef)
	if err != nil {
		return err
	}

	c := s.getCell(col, row)
	if c == nil || len(c.images) == 0 {
		return nil
	}

	for _, img := range c.images {
		if stillUsed := f.imageHrefUsed(img.href, c); !stillUsed {
			delete(f.images, img.href)
		}
	}
	c.images = nil
	return nil
}

func (f *File) imageHrefUsed(href string, excluded *cell) bool {
	for _, s := range f.sheets {
		for _, r := range s.rows {
			for _, c := range r.cells {
				if c == excluded {
					continue
				}
				for _, img := range c.images {
					if img.href == href {
						return true
					}
				}
			}
		}
	}
	return false
}

func detectImageFormat(data []byte) (string, error) {
	switch {
	case len(data) >= 8 && bytes.HasPrefix(data, []byte{0x89, 'P', 'N', 'G', 0x0D, 0x0A, 0x1A, 0x0A}):
		return "png", nil
	case len(data) >= 3 && bytes.HasPrefix(data, []byte{0xFF, 0xD8, 0xFF}):
		return "jpg", nil
	case len(data) >= 6 && (bytes.HasPrefix(data, []byte("GIF87a")) || bytes.HasPrefix(data, []byte("GIF89a"))):
		return "gif", nil
	case len(data) >= 2 && bytes.HasPrefix(data, []byte("BM")):
		return "bmp", nil
	}
	return "", fmt.Errorf("goods: unsupported image format")
}

func imageMimeType(format string) string {
	switch format {
	case "jpg", "jpeg":
		return "image/jpeg"
	case "gif":
		return "image/gif"
	case "bmp":
		return "image/bmp"
	default:
		return "image/png"
	}
}
