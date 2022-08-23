package packaging

import "encoding/xml"

// XCorePropertiesU fix XML ns for XCoreProperties
type XCorePropertiesU struct {
	XMLName       xml.Name `xml:"coreProperties"`
	XmlnsCp       string   `xml:"cp,attr"`
	XmlnsDc       string   `xml:"dc,attr"`
	XmlnsDcterms  string   `xml:"dcterms,attr"`
	XmlnsDcmitype string   `xml:"dcmitype,attr"`
	XmlnsXsi      string   `xml:"xsi,attr"`

	DcCreator        string            `xml:"creator"`
	CpLastModifiedBy string            `xml:"lastModifiedBy"`
	Created          *XDctermsElementU `xml:"created"`
	Modified         *XDctermsElementU `xml:"modified"`
}

// XDctermsElementU fix XML ns for XDctermsElement
type XDctermsElementU struct {
	Text string `xml:",chardata"`
	Type string `xml:"type,attr"`
}
