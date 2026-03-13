package xml

import "encoding/xml"

type Manifest struct {
	XMLName     xml.Name    `xml:"manifest"`
	FileEntries []FileEntry `xml:"file-entry"`
}

type FileEntry struct {
	FullPath  string `xml:"full-path,attr"`
	MediaType string `xml:"media-type,attr"`
	Version   string `xml:"version,attr,omitempty"`
}

func DefaultManifest() Manifest {
	return Manifest{
		FileEntries: []FileEntry{
			{FullPath: "/", MediaType: MimeTypeODS, Version: "1.2"},
			{FullPath: "content.xml", MediaType: "text/xml"},
			{FullPath: "styles.xml", MediaType: "text/xml"},
			{FullPath: "meta.xml", MediaType: "text/xml"},
			{FullPath: "settings.xml", MediaType: "text/xml"},
		},
	}
}
