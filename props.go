package goods

type DocProperties struct {
	Title       string
	Description string
	Subject     string
	Creator     string
}

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
