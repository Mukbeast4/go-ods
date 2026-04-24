package goods

import (
	"bytes"
	"image"
	"image/color"
	"image/png"
	"os"
	"path/filepath"
	"testing"
)

func makeTestPNG(t *testing.T) []byte {
	t.Helper()
	img := image.NewRGBA(image.Rect(0, 0, 2, 2))
	img.Set(0, 0, color.RGBA{R: 255, A: 255})
	img.Set(1, 0, color.RGBA{G: 255, A: 255})
	img.Set(0, 1, color.RGBA{B: 255, A: 255})
	img.Set(1, 1, color.RGBA{R: 255, G: 255, B: 255, A: 255})
	var buf bytes.Buffer
	if err := png.Encode(&buf, img); err != nil {
		t.Fatalf("encode png: %v", err)
	}
	return buf.Bytes()
}

func TestAddImageFromBytes(t *testing.T) {
	f := NewFile()
	data := makeTestPNG(t)

	if err := f.AddImageFromBytes("Sheet1", "B2", data, &ImageOptions{Width: 4, Height: 3}); err != nil {
		t.Fatalf("AddImageFromBytes: %v", err)
	}

	imgs, err := f.GetImages("Sheet1")
	if err != nil {
		t.Fatalf("GetImages: %v", err)
	}
	if len(imgs) != 1 {
		t.Fatalf("want 1 image, got %d", len(imgs))
	}
	if imgs[0].CellRef != "B2" {
		t.Errorf("CellRef = %q, want B2", imgs[0].CellRef)
	}
	if imgs[0].Format != "png" {
		t.Errorf("Format = %q, want png", imgs[0].Format)
	}
	if imgs[0].Width != 4 || imgs[0].Height != 3 {
		t.Errorf("size = %vx%v, want 4x3", imgs[0].Width, imgs[0].Height)
	}
	if !bytes.Equal(imgs[0].Data, data) {
		t.Error("image data mismatch")
	}
}

func TestAddImageRoundTrip(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "roundtrip.ods")

	src := NewFile()
	data := makeTestPNG(t)
	if err := src.AddImageFromBytes("Sheet1", "C5", data, &ImageOptions{Width: 5, Height: 2, OffsetX: 0.5, OffsetY: 0.25}); err != nil {
		t.Fatalf("AddImageFromBytes: %v", err)
	}
	if err := src.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	dst, err := OpenFile(path)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}
	defer dst.Close()

	imgs, err := dst.GetImages("Sheet1")
	if err != nil {
		t.Fatalf("GetImages: %v", err)
	}
	if len(imgs) != 1 {
		t.Fatalf("want 1 image after roundtrip, got %d", len(imgs))
	}
	got := imgs[0]
	if got.CellRef != "C5" {
		t.Errorf("CellRef = %q, want C5", got.CellRef)
	}
	if got.Format != "png" {
		t.Errorf("Format = %q, want png", got.Format)
	}
	if !bytes.Equal(got.Data, data) {
		t.Error("roundtrip data mismatch")
	}
	if got.Width != 5 || got.Height != 2 {
		t.Errorf("size = %vx%v, want 5x2", got.Width, got.Height)
	}
	if got.OffsetX != 0.5 || got.OffsetY != 0.25 {
		t.Errorf("offset = %v,%v, want 0.5,0.25", got.OffsetX, got.OffsetY)
	}
}

func TestAddImageFromFile(t *testing.T) {
	dir := t.TempDir()
	imgPath := filepath.Join(dir, "pic.png")
	data := makeTestPNG(t)
	if err := writeFileImpl(imgPath, data); err != nil {
		t.Fatalf("write png: %v", err)
	}

	f := NewFile()
	if err := f.AddImage("Sheet1", "A1", imgPath, nil); err != nil {
		t.Fatalf("AddImage: %v", err)
	}
	imgs, _ := f.GetImages("Sheet1")
	if len(imgs) != 1 || imgs[0].CellRef != "A1" {
		t.Errorf("unexpected images: %+v", imgs)
	}
}

func writeFileImpl(path string, data []byte) error {
	return os.WriteFile(path, data, 0o644)
}

func TestRemoveImages(t *testing.T) {
	f := NewFile()
	data := makeTestPNG(t)
	f.AddImageFromBytes("Sheet1", "A1", data, nil)
	f.AddImageFromBytes("Sheet1", "B1", data, nil)

	if err := f.RemoveImages("Sheet1", "A1"); err != nil {
		t.Fatalf("RemoveImages: %v", err)
	}

	imgs, _ := f.GetImages("Sheet1")
	if len(imgs) != 1 {
		t.Fatalf("want 1 image remaining, got %d", len(imgs))
	}
	if imgs[0].CellRef != "B1" {
		t.Errorf("remaining image CellRef = %q, want B1", imgs[0].CellRef)
	}
	if len(f.images) != 1 {
		t.Errorf("shared binary should stay while B1 still references it, got %d files", len(f.images))
	}

	f.RemoveImages("Sheet1", "B1")
	if len(f.images) != 0 {
		t.Errorf("binary should be removed when no cell references it, got %d", len(f.images))
	}
}

func TestAddImageInvalidFormat(t *testing.T) {
	f := NewFile()
	err := f.AddImageFromBytes("Sheet1", "A1", []byte("not an image"), nil)
	if err == nil {
		t.Error("want error for invalid format")
	}
}

func TestAddImageSharedBinary(t *testing.T) {
	f := NewFile()
	data := makeTestPNG(t)
	f.AddImageFromBytes("Sheet1", "A1", data, nil)
	f.AddImageFromBytes("Sheet1", "B2", data, nil)

	if len(f.images) != 1 {
		t.Errorf("identical images should share the same binary, got %d", len(f.images))
	}
}

func TestDetectImageFormat(t *testing.T) {
	cases := []struct {
		data []byte
		want string
	}{
		{[]byte{0x89, 'P', 'N', 'G', 0x0D, 0x0A, 0x1A, 0x0A}, "png"},
		{[]byte{0xFF, 0xD8, 0xFF, 0xE0}, "jpg"},
		{[]byte("GIF89a..."), "gif"},
		{[]byte("BM..."), "bmp"},
	}
	for _, tc := range cases {
		got, err := detectImageFormat(tc.data)
		if err != nil {
			t.Errorf("detectImageFormat(%v): %v", tc.data, err)
			continue
		}
		if got != tc.want {
			t.Errorf("detectImageFormat: got %q, want %q", got, tc.want)
		}
	}
}
