package packaging

import "encoding/xml"

// CoreProperties Defines
const (
	CorePropertiesContentType      = "application/vnd.openxmlformats-package.core-properties+xml"
	CorePropertiesRelationshipType = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"

	CorePropertiesPath     = "docProps"
	CorePropertiesFileName = "core.xml"
)

// XCoreProperties directly maps the root element
type XCoreProperties struct {
	XMLName       xml.Name `xml:"cp:coreProperties"`
	XmlnsCp       string   `xml:"xmlns:cp,attr"`
	XmlnsDc       string   `xml:"xmlns:dc,attr"`
	XmlnsDcterms  string   `xml:"xmlns:dcterms,attr"`
	XmlnsDcmitype string   `xml:"xmlns:dcmitype,attr"`
	XmlnsXsi      string   `xml:"xmlns:xsi,attr"`

	DcCreator        string           `xml:"dc:creator"`
	CpLastModifiedBy string           `xml:"cp:lastModifiedBy"`
	Created          *XDctermsElement `xml:"dcterms:created"`
	Modified         *XDctermsElement `xml:"dcterms:modified"`
}

// XDctermsElement document time element
type XDctermsElement struct {
	Text string `xml:",chardata"`
	Type string `xml:"xsi:type,attr"`
}

// NewDefaultXCoreProperties create *XCoreProperties with default template
func NewDefaultXCoreProperties() *XCoreProperties {
	return &XCoreProperties{
		XmlnsCp:       "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
		XmlnsDc:       "http://purl.org/dc/elements/1.1/",
		XmlnsDcterms:  "http://purl.org/dc/terms/",
		XmlnsDcmitype: "http://purl.org/dc/dcmitype/",
		XmlnsXsi:      "http://www.w3.org/2001/XMLSchema-instance",

		DcCreator:        "Microsoft",
		CpLastModifiedBy: "",
		Created: &XDctermsElement{
			Text: "2015-06-05T18:19:34Z",
			Type: "dcterms:W3CDTF",
		},
		Modified: &XDctermsElement{
			Text: "2015-06-05T18:19:39Z",
			Type: "dcterms:W3CDTF",
		},
	}
}
