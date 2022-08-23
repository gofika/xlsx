package packaging

import "encoding/xml"

// XExtendedPropertiesU fix XML ns for XExtendedProperties
type XExtendedPropertiesU struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties Properties"`
	XmlnsVt string   `xml:"vt,attr"`

	Application       string   `xml:"Application"`
	DocSecurity       string   `xml:"DocSecurity"`
	ScaleCrop         string   `xml:"ScaleCrop"`
	HeadingPairs      *XPairsU `xml:"HeadingPairs"`
	TitlesOfParts     *XPairsU `xml:"TitlesOfParts"`
	Company           string   `xml:"Company"`
	LinksUpToDate     string   `xml:"LinksUpToDate"`
	SharedDoc         string   `xml:"SharedDoc"`
	HyperlinksChanged string   `xml:"HyperlinksChanged"`
	AppVersion        string   `xml:"AppVersion"`
}

// XPairsU fix XML ns for XPairs
type XPairsU struct {
	Vector *XVectorU `xml:"vector"`
}

// XVectorU fix XML ns for XVector
type XVectorU struct {
	Size     int    `xml:"size,attr"`
	BaseType string `xml:"baseType,attr"`

	Lpstr   []string     `xml:"lpstr,omitempty"`
	Variant []*XVariantU `xml:"variant,omitempty"`
}

// XVariantU fix XML ns for XVariant
type XVariantU struct {
	Lpstr string `xml:"lpstr,omitempty"`
	I4    int32  `xml:"i4,omitempty"`
}
