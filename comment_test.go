package goods

import (
	"path/filepath"
	"testing"
)

func TestSetGetComment(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	err := f.SetCellComment(s, "A1", &Comment{
		Author: "Alice",
		Date:   "2024-01-15T10:00:00",
		Text:   "This is a comment",
	})
	if err != nil {
		t.Fatalf("SetCellComment: %v", err)
	}

	c, err := f.GetCellComment(s, "A1")
	if err != nil {
		t.Fatalf("GetCellComment: %v", err)
	}
	if c == nil {
		t.Fatal("expected comment, got nil")
	}
	if c.Author != "Alice" {
		t.Errorf("Author = %q, want Alice", c.Author)
	}
	if c.Date != "2024-01-15T10:00:00" {
		t.Errorf("Date = %q, want 2024-01-15T10:00:00", c.Date)
	}
	if c.Text != "This is a comment" {
		t.Errorf("Text = %q, want 'This is a comment'", c.Text)
	}
}

func TestRemoveComment(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellComment(s, "A1", &Comment{Author: "Alice", Text: "test"})
	f.RemoveCellComment(s, "A1")

	c, err := f.GetCellComment(s, "A1")
	if err != nil {
		t.Fatalf("GetCellComment: %v", err)
	}
	if c != nil {
		t.Error("expected nil after remove")
	}
}

func TestCommentNoComment(t *testing.T) {
	f := NewFile()

	c, err := f.GetCellComment("Sheet1", "A1")
	if err != nil {
		t.Fatalf("GetCellComment: %v", err)
	}
	if c != nil {
		t.Error("expected nil for cell without comment")
	}
}

func TestCommentRoundtrip(t *testing.T) {
	f := NewFile()
	s := "Sheet1"

	f.SetCellStr(s, "A1", "Hello")
	f.SetCellComment(s, "A1", &Comment{
		Author: "Bob",
		Date:   "2024-06-01T12:00:00",
		Text:   "Round trip test",
	})

	tmpDir := t.TempDir()
	path := filepath.Join(tmpDir, "comment.ods")

	if err := f.SaveAs(path); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}

	f2, err := OpenFile(path)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}
	defer f2.Close()

	c, err := f2.GetCellComment(s, "A1")
	if err != nil {
		t.Fatalf("GetCellComment: %v", err)
	}
	if c == nil {
		t.Fatal("expected comment after roundtrip")
	}
	if c.Author != "Bob" {
		t.Errorf("Author = %q, want Bob", c.Author)
	}
	if c.Text != "Round trip test" {
		t.Errorf("Text = %q, want 'Round trip test'", c.Text)
	}
}

func TestCommentClosedFile(t *testing.T) {
	f := NewFile()
	f.Close()

	err := f.SetCellComment("Sheet1", "A1", &Comment{Text: "test"})
	if err != ErrFileClosed {
		t.Errorf("expected ErrFileClosed, got %v", err)
	}

	_, err = f.GetCellComment("Sheet1", "A1")
	if err != ErrFileClosed {
		t.Errorf("expected ErrFileClosed, got %v", err)
	}

	err = f.RemoveCellComment("Sheet1", "A1")
	if err != ErrFileClosed {
		t.Errorf("expected ErrFileClosed, got %v", err)
	}
}

func TestCommentSheetNotFound(t *testing.T) {
	f := NewFile()

	err := f.SetCellComment("NoSheet", "A1", &Comment{Text: "test"})
	if err != ErrSheetNotFound {
		t.Errorf("expected ErrSheetNotFound, got %v", err)
	}
}
