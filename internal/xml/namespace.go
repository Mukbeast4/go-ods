package xml

const (
	NsOffice       = "urn:oasis:names:tc:opendocument:xmlns:office:1.0"
	NsStyle        = "urn:oasis:names:tc:opendocument:xmlns:style:1.0"
	NsText         = "urn:oasis:names:tc:opendocument:xmlns:text:1.0"
	NsTable        = "urn:oasis:names:tc:opendocument:xmlns:table:1.0"
	NsDraw         = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
	NsFO           = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
	NsXLink        = "http://www.w3.org/1999/xlink"
	NsDC           = "http://purl.org/dc/elements/1.1/"
	NsMeta         = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0"
	NsNumber       = "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
	NsSVG          = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
	NsChart        = "urn:oasis:names:tc:opendocument:xmlns:chart:1.0"
	NsDR3D         = "urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0"
	NsForm         = "urn:oasis:names:tc:opendocument:xmlns:form:1.0"
	NsScript       = "urn:oasis:names:tc:opendocument:xmlns:script:1.0"
	NsPresentation = "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"
	NsManifest     = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"
	NsCalcExt      = "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0"
	NsLoExt        = "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"
	NsFieldD       = "urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0"
	NsCSS3T        = "http://www.w3.org/TR/css3-text/"

	MimeTypeODS = "application/vnd.oasis.opendocument.spreadsheet"
)

var PrefixToNS = map[string]string{
	"office":       NsOffice,
	"style":        NsStyle,
	"text":         NsText,
	"table":        NsTable,
	"draw":         NsDraw,
	"fo":           NsFO,
	"xlink":        NsXLink,
	"dc":           NsDC,
	"meta":         NsMeta,
	"number":       NsNumber,
	"svg":          NsSVG,
	"chart":        NsChart,
	"dr3d":         NsDR3D,
	"form":         NsForm,
	"script":       NsScript,
	"presentation": NsPresentation,
	"manifest":     NsManifest,
	"calcext":      NsCalcExt,
	"loext":        NsLoExt,
	"field":        NsFieldD,
	"css3t":        NsCSS3T,
}
