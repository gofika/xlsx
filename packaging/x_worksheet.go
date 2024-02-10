package packaging

import (
	"encoding/xml"

	"github.com/shopspring/decimal"
)

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
	MergeCells    *XMergeCells    `xml:"mergeCells,omitempty"`
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
	DefaultRowHeight decimal.Decimal `xml:"defaultRowHeight,attr"`
	X14acDyDescent   decimal.Decimal `xml:"x14ac:dyDescent,attr"`
}

// XSheetData SheetData node
type XSheetData struct {
	XMLName xml.Name `xml:"sheetData"`
	Row     []*XRow  `xml:"row"`
}

// XRow Row node
type XRow struct {
	R              int             `xml:"r,attr"` // row number
	Spans          string          `xml:"spans,attr,omitempty"`
	Hidden         BoolAttr        `xml:"hidden,attr,omitempty"`
	C              []*XC           `xml:"c"`
	Ht             string          `xml:"ht,attr,omitempty"`
	CustomHeight   BoolAttr        `xml:"customHeight,attr,omitempty"`
	OutlineLevel   uint8           `xml:"outlineLevel,attr,omitempty"`
	S              int             `xml:"s,attr,omitempty"`            // row style id
	CustomFormat   BoolAttr        `xml:"customFormat,attr,omitempty"` // enable row custom format
	X14acDyDescent decimal.Decimal `xml:"x14ac:dyDescent,attr"`
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
	Content string      `xml:",chardata"`
	T       OmitIntAttr `xml:"t,attr,omitempty"`   // Formula type
	Ref     string      `xml:"ref,attr,omitempty"` // Shared formula ref
	Si      int         `xml:"si,attr,omitempty"`  // Shared formula index
}

// XPageMargins PageMargins node
type XPageMargins struct {
	Left   decimal.Decimal `xml:"left,attr"`
	Right  decimal.Decimal `xml:"right,attr"`
	Top    decimal.Decimal `xml:"top,attr"`
	Bottom decimal.Decimal `xml:"bottom,attr"`
	Header decimal.Decimal `xml:"header,attr"`
	Footer decimal.Decimal `xml:"footer,attr"`
}

// XCols Cols node
type XCols struct {
	Col []*XCol `xml:"col"`
}

// XCol Col node
type XCol struct {
	Min          int             `xml:"min,attr"`
	Max          int             `xml:"max,attr"`
	BestFit      BoolAttr        `xml:"bestFit,attr,omitempty"`
	Collapsed    BoolAttr        `xml:"collapsed,attr,omitempty"`
	CustomWidth  BoolAttr        `xml:"customWidth,attr,omitempty"`
	Hidden       BoolAttr        `xml:"hidden,attr,omitempty"`
	OutlineLevel uint8           `xml:"outlineLevel,attr,omitempty"`
	Phonetic     BoolAttr        `xml:"phonetic,attr,omitempty"`
	Style        int             `xml:"style,attr,omitempty"`
	Width        decimal.Decimal `xml:"width,attr,omitempty"`
}

// XMergeCells MergeCells node
type XMergeCells struct {
	Count     int           `xml:"count,attr,omitempty"`
	MergeCell []*XMergeCell `xml:"mergeCell,omitempty"`
}

// XMergeCell MergeCell node
type XMergeCell struct {
	Ref string `xml:"ref,attr,omitempty"`
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
			DefaultRowHeight: decimal.NewFromInt(15),
			X14acDyDescent:   decimal.NewFromFloat(0.25),
		},
		SheetData: &XSheetData{},
		PageMargins: &XPageMargins{
			Left:   decimal.NewFromFloat(0.7),
			Right:  decimal.NewFromFloat(0.7),
			Top:    decimal.NewFromFloat(0.75),
			Bottom: decimal.NewFromFloat(0.75),
			Header: decimal.NewFromFloat(0.3),
			Footer: decimal.NewFromFloat(0.3),
		},
	}
}
