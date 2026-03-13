package xml

import "encoding/xml"

type DocumentMeta struct {
	XMLName xml.Name `xml:"document-meta"`
	Meta    Meta     `xml:"meta"`
}

type Meta struct {
	Generator      string `xml:"generator,omitempty"`
	Title          string `xml:"title,omitempty"`
	Description    string `xml:"description,omitempty"`
	Subject        string `xml:"subject,omitempty"`
	InitialCreator string `xml:"initial-creator,omitempty"`
	Creator        string `xml:"creator,omitempty"`
	CreationDate   string `xml:"creation-date,omitempty"`
	Date           string `xml:"date,omitempty"`
}
