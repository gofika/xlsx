package packaging

import "encoding/xml"

//  SharedStrings Defines
const (
	SharedStringsContentType      = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
	SharedStringsRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"

	SharedStringsPath     = "xl"
	SharedStringsFileName = "sharedStrings.xml"
)

// XSharedStrings XML document
type XSharedStrings struct {
	XMLName     xml.Name `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main sst"`
	Count       int      `xml:"count,attr"`
	UniqueCount int      `xml:"uniqueCount,attr"`

	Si []*XSi `xml:"si,omitempty"`
}

// XSi Si node
type XSi struct {
	T          string       `xml:"t"`
	PhoneticPr *XPhoneticPr `xml:"phoneticPr"`
}

// XPhoneticPr PhoneticPr node
type XPhoneticPr struct {
	FontID string `xml:"fontId,attr"`
	Type   string `xml:"type,attr"`
}

// NewDefaultXSharedStrings create *XSharedStrings with default template
func NewDefaultXSharedStrings() *XSharedStrings {
	return &XSharedStrings{
		Count:       0,
		UniqueCount: 0,
		Si:          []*XSi{},
	}
}
