package packaging

import "encoding/xml"

// XStyleSheetU fix XML ns for XStyleSheet
type XStyleSheetU struct {
	XMLName     xml.Name `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main styleSheet"`
	XmlnsMc     string   `xml:"mc,attr"`
	McIgnorable string   `xml:"Ignorable,attr"`
	XmlnsX14ac  string   `xml:"x14ac,attr"`
	XmlnsX16r2  string   `xml:"x16r2,attr"`

	Fonts        *XFontsU            `xml:"fonts"`
	Fills        *XFillsU            `xml:"fills"`
	Borders      *XBordersU          `xml:"borders"`
	CellStyleXfs *XCellStyleXfsU     `xml:"cellStyleXfs"`
	CellXfs      *XCellXfsU          `xml:"cellXfs"`
	CellStyles   *XCellStylesU       `xml:"cellStyles"`
	Dxfs         *XDxfsU             `xml:"dxfs"`
	TableStyles  *XTableStylesU      `xml:"tableStyles"`
	ExtLst       *XStyleSheetExtLstU `xml:"extLst"`
	NumFmts      *XNumFmtsU          `xml:"numFmts,omitempty"`
}

// XFontsU fix XML ns for XFonts
type XFontsU struct {
	Count      int    `xml:"count,attr"`
	KnownFonts string `xml:"knownFonts,attr"`

	Font []*XStyleSheetFontU `xml:"font"`
}

// XStyleSheetFontU fix XML ns for XStyleSheetFont
type XStyleSheetFontU struct {
	Sz      *XIntValAttrU     `xml:"sz"`
	Color   *XColorU          `xml:"color"`
	Name    *XValAttrElementU `xml:"name"`
	Family  *XValAttrElementU `xml:"family"`
	Scheme  *XValAttrElementU `xml:"scheme"`
	B       *XBoolValAttrU    `xml:"b,omitempty"`
	I       *XBoolValAttrU    `xml:"i,omitempty"`
	U       *XValAttrElementU `xml:"u,omitempty"`
	Strike  *XValAttrElementU `xml:"strike,omitempty"`
	Charset *XValAttrElement  `xml:"charset,omitempty"` // Character Set
}

// XIntValAttrU fix XML ns for XIntValAttr
type XIntValAttrU struct {
	Val int `xml:"val,attr"`
}

// XBoolValAttrU fix XML ns for XBoolValAttr
type XBoolValAttrU struct {
	Val bool `xml:"val,attr"`
}

// XColorU fix XML ns for XColor
type XColorU struct {
	Theme   string `xml:"theme,attr"`
	RGB     string `xml:"rgb,attr,omitempty"`
	Auto    bool   `xml:"auto,attr,omitempty"`
	Indexed string `xml:"indexed,attr,omitempty"`
	Tint    string `xml:"tint,attr,omitempty"`
}

// XFillsU fix XML ns for XFills
type XFillsU struct {
	Count int `xml:"count,attr"`

	Fill []*XFillU `xml:"fill"`
}

// XFillU fix XML ns for XFill
type XFillU struct {
	PatternFill *XPatternFillU `xml:"patternFill"`
}

// XPatternFillU fix XML ns for XPatternFill
type XPatternFillU struct {
	PatternType string `xml:"patternType,attr"`
}

// XBordersU fix XML ns for XBorders
type XBordersU struct {
	Count int `xml:"count,attr"`

	Border []*XBorderU `xml:"border"`
}

// XBorderU fix XML ns for XBorder
type XBorderU struct {
	Left     string `xml:"left"`
	Right    string `xml:"right"`
	Top      string `xml:"top"`
	Bottom   string `xml:"bottom"`
	Diagonal string `xml:"diagonal"`
}

// XCellStyleXfsU fix XML ns for XCellStyleXfs
type XCellStyleXfsU struct {
	Count int `xml:"count,attr"`

	Xf []*XXfU `xml:"xf"`
}

// XXfU fix XML ns for XXf
type XXfU struct {
	NumFmtID          int  `xml:"numFmtId,attr"`
	FontID            int  `xml:"fontId,attr"`
	FillID            int  `xml:"fillId,attr"`
	BorderID          int  `xml:"borderId,attr"`
	XfID              int  `xml:"xfId,attr,omitempty"`
	ApplyAlignment    bool `xml:"applyAlignment,attr,omitempty"`
	ApplyBorder       bool `xml:"applyBorder,attr,omitempty"`
	ApplyFill         bool `xml:"applyFill,attr,omitempty"`
	ApplyFont         bool `xml:"applyFont,attr,omitempty"`
	ApplyNumberFormat bool `xml:"applyNumberFormat,attr,omitempty"`
	ApplyProtection   bool `xml:"applyProtection,attr,omitempty"`
	PivotButton       bool `xml:"pivotButton,attr,omitempty"`
	QuotePrefix       bool `xml:"quotePrefix,attr,omitempty"`
}

// XCellXfsU fix XML ns for XCellXfs
type XCellXfsU struct {
	Count int `xml:"count,attr"`

	Xf []*XXfU `xml:"xf"`
}

// XCellStylesU fix XML ns for XCellStyles
type XCellStylesU struct {
	Count int `xml:"count,attr"`

	CellStyle []*XCellStyleU `xml:"cellStyle"`
}

// XCellStyleU fix XML ns for XCellStyle
type XCellStyleU struct {
	Name          string `xml:"name,attr"`
	XfID          int    `xml:"xfId,attr"`
	BuiltinID     int    `xml:"builtinId,attr"`
	CustomBuiltin string `xml:"customBuiltin,attr,omitempty"`
	Hidden        string `xml:"hidden,attr,omitempty"`
	ILevel        string `xml:"iLevel,attr,omitempty"`
}

// XDxfsU fix XML ns for XDxfs
type XDxfsU struct {
	Count int `xml:"count,attr"`
}

// XTableStylesU fix XML ns for XTableStyles
type XTableStylesU struct {
	Count             int    `xml:"count,attr"`
	DefaultTableStyle string `xml:"defaultTableStyle,attr"`
	DefaultPivotStyle string `xml:"defaultPivotStyle,attr"`
}

// XStyleSheetExtLstU fix XML ns for XStyleSheetExtLst
type XStyleSheetExtLstU struct {
	Ext []*XStyleSheetExtU `xml:"ext"`
}

// XStyleSheetExtU fix XML ns for XStyleSheetExt
type XStyleSheetExtU struct {
	URI            string            `xml:"uri,attr"`
	XmlnsX14       string            `xml:"x14,attr,omitempty"`
	XmlnsX15       string            `xml:"x15,attr,omitempty"`
	SlicerStyles   *XSlicerStylesU   `xml:"slicerStyles,omitempty"`
	TimelineStyles *XTimelineStylesU `xml:"timelineStyles,omitempty"`
}

// XSlicerStylesU fix XML ns for XSlicerStyles
type XSlicerStylesU struct {
	DefaultSlicerStyle string `xml:"defaultSlicerStyle,attr"`
}

// XTimelineStylesU fix XML ns for XTimelineStyles
type XTimelineStylesU struct {
	DefaultTimelineStyle string `xml:"defaultTimelineStyle,attr"`
}

// XNumFmtsU fix XML ns for XNumFmts
type XNumFmtsU struct {
	Count  int         `xml:"count,attr"`
	NumFmt []*XNumFmtU `xml:"numFmt"`
}

// XNumFmtU fix XML ns for XNumFmt
type XNumFmtU struct {
	FormatCode string `xml:"formatCode,attr"`
	NumFmtId   int    `xml:"numFmtId,attr"`
}
