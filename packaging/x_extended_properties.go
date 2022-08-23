package packaging

import "encoding/xml"

// ExtendedProperties Defines
const (
	ExtendedPropertiesContentType      = "application/vnd.openxmlformats-officedocument.extended-properties+xml"
	ExtendedPropertiesRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"

	ExtendedPropertiesPath     = "docProps"
	ExtendedPropertiesFileName = "app.xml"
)

// XExtendedProperties Document extended properties
type XExtendedProperties struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties Properties"`
	XmlnsVt string   `xml:"xmlns:vt,attr"`

	Application       string  `xml:"Application"`
	DocSecurity       string  `xml:"DocSecurity"`
	ScaleCrop         string  `xml:"ScaleCrop"`
	HeadingPairs      *XPairs `xml:"HeadingPairs"`
	TitlesOfParts     *XPairs `xml:"TitlesOfParts"`
	Company           string  `xml:"Company"`
	LinksUpToDate     string  `xml:"LinksUpToDate"`
	SharedDoc         string  `xml:"SharedDoc"`
	HyperlinksChanged string  `xml:"HyperlinksChanged"`
	AppVersion        string  `xml:"AppVersion"`
}

// XPairs Pairs Type
type XPairs struct {
	Vector *XVector `xml:"vt:vector"`
}

// XVector Vector Type
type XVector struct {
	Size     int    `xml:"size,attr"`
	BaseType string `xml:"baseType,attr"`

	Lpstr   []string    `xml:"vt:lpstr,omitempty"`
	Variant []*XVariant `xml:"vt:variant,omitempty"`
}

// XVariant Variant Type
type XVariant struct {
	Lpstr string `xml:"vt:lpstr,omitempty"`
	I4    int32  `xml:"vt:i4,omitempty"`
}

// NewXExtendedProperties create *XExtendedProperties from Workbook
func NewXExtendedProperties(workbook *XWorkbook) (properties *XExtendedProperties) {
	properties = &XExtendedProperties{
		XmlnsVt: "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",

		Application: "Microsoft Excel",
		DocSecurity: "0",
		ScaleCrop:   "false",
		HeadingPairs: &XPairs{
			Vector: &XVector{
				Size:     2,
				BaseType: "variant",
				Variant: []*XVariant{
					{
						Lpstr: "Worksheets",
					},
					{
						I4: int32(len(workbook.Sheets.Sheet)),
					},
				},
			},
		},
		TitlesOfParts: &XPairs{
			Vector: &XVector{
				Size:     len(workbook.Sheets.Sheet),
				BaseType: "lpstr",
			},
		},
		Company:           "",
		LinksUpToDate:     "false",
		SharedDoc:         "false",
		HyperlinksChanged: "false",
		AppVersion:        "16.0300",
	}

	for _, sheet := range workbook.Sheets.Sheet {
		properties.TitlesOfParts.Vector.Lpstr = append(properties.TitlesOfParts.Vector.Lpstr, sheet.Name)
	}

	return
}
