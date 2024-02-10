package packaging

import (
	"encoding/xml"

	"github.com/shopspring/decimal"
)

// XStyleSheetU fix XML ns for XStyleSheet
type XStyleSheetU struct {
	XMLName     xml.Name `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main styleSheet"`
	XmlnsMc     string   `xml:"mc,attr"`
	McIgnorable string   `xml:"Ignorable,attr"`
	XmlnsX14ac  string   `xml:"x14ac,attr"`
	XmlnsX16r2  string   `xml:"x16r2,attr"`

	NumFmts      *XNumFmtsU          `xml:"numFmts,omitempty"`
	Fonts        *XFontsU            `xml:"fonts"`
	Fills        *XFillsU            `xml:"fills"`
	Borders      *XBordersU          `xml:"borders"`
	CellStyleXfs *XCellStyleXfsU     `xml:"cellStyleXfs"`
	CellXfs      *XCellXfsU          `xml:"cellXfs"`
	CellStyles   *XCellStylesU       `xml:"cellStyles"`
	Dxfs         *XDxfsU             `xml:"dxfs"`
	TableStyles  *XTableStylesU      `xml:"tableStyles"`
	ExtLst       *XStyleSheetExtLstU `xml:"extLst"`
}

// XFontsU fix XML ns for XFonts
type XFontsU struct {
	Count      int    `xml:"count,attr"`
	KnownFonts string `xml:"knownFonts,attr"`

	Font []*XStyleSheetFontU `xml:"font"`
}

// XStyleSheetFontU fix XML ns for XStyleSheetFont
type XStyleSheetFontU struct {
	B         *XBoolValAttr    `xml:"b,omitempty"`         // Bold
	Charset   *XIntValAttr     `xml:"charset,omitempty"`   // Character Set
	Color     *XColor          `xml:"color,omitempty"`     // Text Color
	Condense  *XBoolValAttr    `xml:"condense,omitempty"`  // Condense
	Extend    *XBoolValAttr    `xml:"extend,omitempty"`    // Extend
	Family    *XIntValAttr     `xml:"family"`              // Font Family
	I         *XBoolValAttr    `xml:"i,omitempty"`         // Italic
	Name      *XValAttrElement `xml:"name"`                // Font Name
	Outline   *XBoolValAttr    `xml:"outline,omitempty"`   // Outline
	Scheme    *XValAttrElement `xml:"scheme,omitempty"`    // Scheme
	Shadow    *XBoolValAttr    `xml:"shadow,omitempty"`    // Shadow
	Strike    *XBoolValAttr    `xml:"strike,omitempty"`    // Strike Through. TextStrikeValues
	Sz        *XIntValAttr     `xml:"sz"`                  // Font Size
	U         *XValAttrElement `xml:"u,omitempty"`         // Underline. TextUnderlineValues
	VertAlign *XValAttrElement `xml:"vertAlign,omitempty"` // Vertical Alignment
}

// XIntValAttrU fix XML ns for XIntValAttr
type XIntValAttrU struct {
	Val int `xml:"val,attr"`
}

// XBoolValAttrU fix XML ns for XBoolValAttr
type XBoolValAttrU struct {
	Val BoolAttr `xml:"val,attr"`
}

// XColorU fix XML ns for XColor
type XColorU struct {
	Theme   OmitIntAttr     `xml:"theme,attr"`
	RGB     string          `xml:"rgb,attr,omitempty"`
	Auto    BoolAttr        `xml:"auto,attr,omitempty"`
	Indexed OmitIntAttr     `xml:"indexed,attr,omitempty"`
	Tint    decimal.Decimal `xml:"tint,attr,omitempty"`
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
	PatternType     string  `xml:"patternType,attr"`
	BackgroundColor *XColor `xml:"bgColor,omitempty"`
	ForegroundColor *XColor `xml:"fgColor,omitempty"`
}

// XBordersU fix XML ns for XBorders
type XBordersU struct {
	Count int `xml:"count,attr"`

	Border []*XBorder `xml:"border"`
}

// XCellStyleXfsU fix XML ns for XCellStyleXfs
type XCellStyleXfsU struct {
	Count int `xml:"count,attr"`

	Xf []*XXf `xml:"xf"`
}

// XCellXfsU fix XML ns for XCellXfs
type XCellXfsU struct {
	Count int `xml:"count,attr"`

	Xf []*XXf `xml:"xf"`
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
	NumFmtID   int    `xml:"numFmtId,attr"`
}
