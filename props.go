package goods

// DocProperties holds document metadata written to meta.xml. Empty fields are
// ignored by SetDocProperties so existing values are not overwritten.
type DocProperties struct {
	Title       string
	Description string
	Subject     string
	Creator     string
}

// SetDocProperties merges props into the current document metadata. Empty
// fields in props are left untouched; use the dedicated getters if you need to
// clear a field.
func (f *File) SetDocProperties(props *DocProperties) error {
	if f.closed {
		return ErrFileClosed
	}

	if props.Title != "" {
		f.metadata.Meta.Title = props.Title
	}
	if props.Description != "" {
		f.metadata.Meta.Description = props.Description
	}
	if props.Subject != "" {
		f.metadata.Meta.Subject = props.Subject
	}
	if props.Creator != "" {
		f.metadata.Meta.Creator = props.Creator
	}

	return nil
}

// GetDocProperties returns a copy of the current document metadata.
func (f *File) GetDocProperties() (*DocProperties, error) {
	if f.closed {
		return nil, ErrFileClosed
	}

	return &DocProperties{
		Title:       f.metadata.Meta.Title,
		Description: f.metadata.Meta.Description,
		Subject:     f.metadata.Meta.Subject,
		Creator:     f.metadata.Meta.Creator,
	}, nil
}
