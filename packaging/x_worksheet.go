package packaging

import "encoding/xml"

// Worksheet Defines
const (
	WorksheetContentType      = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
	WorksheetRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"

	WorksheetPath     = "xl"
	WorksheetFileName = "worksheets/sheet%d.xml"
)

// XWorksheet Worksheet XML doc
type XWorksheet struct {
	XMLName     xml.Name `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main worksheet"`
	XmlnsR      string   `xml:"xmlns:r,attr"`
	XmlnsMc     string   `xml:"xmlns:mc,attr"`
	XmlnsX14ac  string   `xml:"xmlns:x14ac,attr"`
	McIgnorable string   `xml:"mc:Ignorable,attr"`

	Dimension     *XDimension     `xml:"dimension"`
	SheetViews    *XSheetViews    `xml:"sheetViews"`
	SheetFormatPr *XSheetFormatPr `xml:"sheetFormatPr"`
	Cols          *XCols          `xml:"cols,omitempty"`
	SheetData     *XSheetData     `xml:"sheetData"`
	PageMargins   *XPageMargins   `xml:"pageMargins"`
}

// XDimension Dimension node
type XDimension struct {
	Ref string `xml:"ref,attr"`
}

// XSheetViews SheetViews node
type XSheetViews struct {
	SheetView *XSheetView `xml:"sheetView"`
}

// XSheetView SheetView node
type XSheetView struct {
	TabSelected    int `xml:"tabSelected,attr,omitempty"`
	WorkbookViewID int `xml:"workbookViewId,attr"`
}

// XSheetFormatPr SheetFormatPr node
type XSheetFormatPr struct {
	DefaultRowHeight float64 `xml:"defaultRowHeight,attr"`
	X14acDyDescent   float64 `xml:"x14ac:dyDescent,attr"`
}

// XSheetData SheetData node
type XSheetData struct {
	XMLName xml.Name `xml:"sheetData"`
	Row     []*XRow  `xml:"row"`
}

// XRow Row node
type XRow struct {
	R              int     `xml:"r,attr"` // row number
	Spans          string  `xml:"spans,attr,omitempty"`
	Hidden         bool    `xml:"hidden,attr,omitempty"`
	C              []*XC   `xml:"c"`
	Ht             string  `xml:"ht,attr,omitempty"`
	CustomHeight   bool    `xml:"customHeight,attr,omitempty"`
	OutlineLevel   uint8   `xml:"outlineLevel,attr,omitempty"`
	S              int     `xml:"s,attr,omitempty"`            // row style id
	CustomFormat   bool    `xml:"customFormat,attr,omitempty"` // enable row custom format
	X14acDyDescent float64 `xml:"x14ac:dyDescent,attr"`
}

// XC This collection represents a cell in the worksheet. Information about the cell's location (reference), value, data type, formatting, and formula is expressed here.
type XC struct {
	XMLName xml.Name

	R string `xml:"r,attr"`
	S int    `xml:"s,attr,omitempty"`
	T string `xml:"t,attr,omitempty"`
	F *XF    `xml:"f,omitempty"`
	V string `xml:"v,omitempty"`
}

// XF F node
type XF struct {
	Content string `xml:",chardata"`
	T       string `xml:"t,attr,omitempty"`   // Formula type
	Ref     string `xml:"ref,attr,omitempty"` // Shared formula ref
	Si      int    `xml:"si,attr,omitempty"`  // Shared formula index
}

// XPageMargins PageMargins node
type XPageMargins struct {
	Left   float64 `xml:"left,attr"`
	Right  float64 `xml:"right,attr"`
	Top    float64 `xml:"top,attr"`
	Bottom float64 `xml:"bottom,attr"`
	Header float64 `xml:"header,attr"`
	Footer float64 `xml:"footer,attr"`
}

// XCols Cols node
type XCols struct {
	Col []*XCol `xml:"col"`
}

// XCol Col node
type XCol struct {
	Min          int     `xml:"min,attr"`
	Max          int     `xml:"max,attr"`
	BestFit      bool    `xml:"bestFit,attr,omitempty"`
	Collapsed    bool    `xml:"collapsed,attr,omitempty"`
	CustomWidth  bool    `xml:"customWidth,attr,omitempty"`
	Hidden       bool    `xml:"hidden,attr,omitempty"`
	OutlineLevel uint8   `xml:"outlineLevel,attr,omitempty"`
	Phonetic     bool    `xml:"phonetic,attr,omitempty"`
	Style        int     `xml:"style,attr,omitempty"`
	Width        float64 `xml:"width,attr,omitempty"`
}

// NewDefaultXWorksheet create *XWorksheet with default template
func NewDefaultXWorksheet() *XWorksheet {
	return &XWorksheet{
		XmlnsR:      "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
		XmlnsMc:     "http://schemas.openxmlformats.org/markup-compatibility/2006",
		XmlnsX14ac:  "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
		McIgnorable: "x14ac",

		Dimension: &XDimension{
			Ref: "A1",
		},
		SheetViews: &XSheetViews{
			SheetView: &XSheetView{
				TabSelected:    1,
				WorkbookViewID: 0,
			},
		},
		SheetFormatPr: &XSheetFormatPr{
			DefaultRowHeight: 15,
			X14acDyDescent:   0.25,
		},
		SheetData: &XSheetData{},
		PageMargins: &XPageMargins{
			Left:   0.7,
			Right:  0.7,
			Top:    0.75,
			Bottom: 0.75,
			Header: 0.3,
			Footer: 0.3,
		},
	}
}
