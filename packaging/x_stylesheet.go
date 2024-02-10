package packaging

import (
	"encoding/xml"
)

// StyleSheet Defines
const (
	StyleSheetContentType      = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
	StyleSheetRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"

	StyleSheetPath     = "xl"
	StyleSheetFileName = "styles.xml"
)

// TextUnderlineValues
const (
	TextUnderlineNone            = "none"
	TextUnderlineWords           = "words"
	TextUnderlineSingle          = "sng"
	TextUnderlineDouble          = "dbl"
	TextUnderlineHeavy           = "heavy"
	TextUnderlineDotted          = "dotted"
	TextUnderlineHeavyDotted     = "dottedHeavy"
	TextUnderlineDash            = "dash"
	TextUnderlineDashHeavy       = "dashHeavy"
	TextUnderlineDashLong        = "dashLong"
	TextUnderlineDashLongHeavy   = "dashLongHeavy"
	TextUnderlineDotDash         = "dotDash"
	TextUnderlineDotDashHeavy    = "dotDashHeavy"
	TextUnderlineDotDotDash      = "dotDotDash"
	TextUnderlineDotDotDashHeavy = "dotDotDashHeavy"
	TextUnderlineWavy            = "wavy"
	TextUnderlineWavyHeavy       = "wavyHeavy"
	TextUnderlineWavyDouble      = "wavyDbl"
)

// TextStrikeValues
const (
	TextStrikeNoStrike     = "noStrike"
	TextStrikeSingleStrike = "sngStrike"
	TextStrikeDoubleStrike = "dblStrike"
)

// XStyleSheet StyleSheet XML document
type XStyleSheet struct {
	XMLName     xml.Name `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main styleSheet"`
	XmlnsMc     string   `xml:"xmlns:mc,attr"`
	McIgnorable string   `xml:"mc:Ignorable,attr"`
	XmlnsX14ac  string   `xml:"xmlns:x14ac,attr"`
	XmlnsX16r2  string   `xml:"xmlns:x16r2,attr"`

	NumFmts      *XNumFmts          `xml:"numFmts,omitempty"` // Number Formats
	Fonts        *XFonts            `xml:"fonts"`             // Fonts
	Fills        *XFills            `xml:"fills"`             // Fills
	Borders      *XBorders          `xml:"borders"`           // Borders
	CellStyleXfs *XCellStyleXfs     `xml:"cellStyleXfs"`      // Formatting Records
	CellXfs      *XCellXfs          `xml:"cellXfs"`           // Cell Formats
	CellStyles   *XCellStyles       `xml:"cellStyles"`        // Cell Styles
	Dxfs         *XDxfs             `xml:"dxfs"`              // Formats
	TableStyles  *XTableStyles      `xml:"tableStyles"`       // Table Styles
	ExtLst       *XStyleSheetExtLst `xml:"extLst"`            // Future Feature Data Storage Area
	//Colors       *XColors           `xml:"x:colors"` // Colors
}

// XFonts Fonts type
type XFonts struct {
	Count      int    `xml:"count,attr"`
	KnownFonts string `xml:"x14ac:knownFonts,attr"`

	Font []*XStyleSheetFont `xml:"font"`
}

// XStyleSheetFont This element defines the properties for one of the fonts used in this workbook.
// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.font?view=openxml-2.8.1
type XStyleSheetFont struct {
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

func (f *XStyleSheetFont) Equal(other *XStyleSheetFont) bool {
	if f == nil && other == nil {
		return true
	}
	if f == nil || other == nil {
		return false
	}
	return f.B.Equal(other.B) &&
		f.Charset.Equal(other.Charset) &&
		f.Condense.Equal(other.Condense) &&
		f.Extend.Equal(other.Extend) &&
		f.Family.Equal(other.Family) &&
		f.I.Equal(other.I) &&
		f.Name.Equal(other.Name) &&
		f.Outline.Equal(other.Outline) &&
		f.Scheme.Equal(other.Scheme) &&
		f.Shadow.Equal(other.Shadow) &&
		f.Strike.Equal(other.Strike) &&
		f.Sz.Equal(other.Sz) &&
		f.U.Equal(other.U) &&
		f.VertAlign.Equal(other.VertAlign)
}

// XValAttrElement Val type
type XValAttrElement struct {
	Val string `xml:"val,attr"`
}

func (a *XValAttrElement) Value() string {
	if a == nil {
		return ""
	}
	return a.Val
}

func (a *XValAttrElement) Equal(b *XValAttrElement) bool {
	return a.Value() == b.Value()
}

// XIntValAttr
// Example: <sz val="11" />
type XIntValAttr struct {
	Val int `xml:"val,attr"`
}

func (a *XIntValAttr) Value() int {
	if a == nil {
		return 0
	}
	return a.Val
}

func (a *XIntValAttr) Equal(other *XIntValAttr) bool {
	return a.Value() == other.Value()
}

// XBoolValAttr
// Example: <b val="1" />
type XBoolValAttr struct {
	Val BoolAttr `xml:"val,attr"`
}

func (a *XBoolValAttr) Value() bool {
	if a == nil {
		return false
	}
	return a.Val.Value()
}

func (a *XBoolValAttr) Equal(other *XBoolValAttr) bool {
	return a.Value() == other.Value()
}

// XColor Data Bar Color
// One of the colors associated with the data bar or color scale.
// The auto attribute shall not be used in the context of data bars.
type XColor struct {
	// Theme Color. A zero-based index into the <clrScheme> collection (§20.1.6.2), referencing a particular <sysClr> or <srgbClr> value expressed in the Theme part.
	Theme OmitIntAttr `xml:"theme,attr,omitempty"`
	// Alpha Red Green Blue Color Value. Standard Alpha Red Green Blue color value (ARGB).
	RGB string `xml:"rgb,attr,omitempty"`
	// Automatic. A boolean value indicating the color is automatic and system color dependent.
	Auto BoolAttr `xml:"auto,attr,omitempty"`
	// Index. Indexed color value. Only used for backwards compatibility. References a color in indexedColors.
	Indexed OmitIntAttr `xml:"indexed,attr,omitempty"`
	// Tint. Specifies the tint value applied to the color.
	// If tint is supplied, then it is applied to the value of the color to determine the final color applied.
	// The tint value is stored as a double from -1.0 .. 1.0, where -1.0 means 100% darken and 1.0 means 100% lighten. Also, 0.0 means no change.
	Tint string `xml:"tint,attr,omitempty"`
}

func (c *XColor) Equal(other *XColor) bool {
	if c == nil && other == nil {
		return true
	}
	var self *XColor
	if c == nil {
		self = &XColor{}
	} else {
		self = c
	}
	if other == nil {
		other = &XColor{}
	}
	return self.Theme == other.Theme &&
		self.RGB == other.RGB &&
		self.Auto == other.Auto &&
		self.Indexed == other.Indexed &&
		self.Tint == other.Tint
}

// XFills Fills type
type XFills struct {
	Count int `xml:"count,attr"`

	Fill []*XFill `xml:"fill"`
}

// XFill Fill type
type XFill struct {
	PatternFill *XPatternFill `xml:"patternFill"`
}

func (f *XFill) GetPatternType() string {
	if f == nil || f.PatternFill == nil {
		return ""
	}
	return f.PatternFill.PatternType
}

func (f *XFill) Equal(other *XFill) bool {
	if f == nil && other == nil {
		return true
	}
	var self *XFill
	if f == nil {
		self = &XFill{}
	} else {
		self = f
	}
	if other == nil {
		other = &XFill{}
	}
	return self.GetPatternType() == other.GetPatternType()
}

// XPatternFill PatternFill type
type XPatternFill struct {
	PatternType     string  `xml:"patternType,attr"`
	BackgroundColor *XColor `xml:"bgColor,omitempty"`
	ForegroundColor *XColor `xml:"fgColor,omitempty"`
}

// XBorders Borders type
type XBorders struct {
	Count  int        `xml:"count,attr"`
	Border []*XBorder `xml:"border"`
}

// XBorder Border type
// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.border?view=openxml-2.8.1
type XBorder struct {
	DiagonalDown BoolAttr     `xml:"diagonalDown,attr,omitempty"` // A boolean value indicating if the cell's diagonal border includes a diagonal line, starting at the top left corner of the cell and moving down to the bottom right corner of the cell.
	DiagonalUp   BoolAttr     `xml:"diagonalUp,attr,omitempty"`   // A boolean value indicating if the cell's diagonal border includes a diagonal line, starting at the bottom left corner of the cell and moving up to the top right corner of the cell.
	Outline      BoolAttr     `xml:"outline,attr,omitempty"`      // A boolean value indicating if left, right, top, and bottom borders should be applied only to outside borders of a cell range.
	Left         *XBorderInfo `xml:"left,omitempty"`
	Right        *XBorderInfo `xml:"right,omitempty"`
	Top          *XBorderInfo `xml:"top,omitempty"`
	Bottom       *XBorderInfo `xml:"bottom,omitempty"`
	Diagonal     *XBorderInfo `xml:"diagonal,omitempty"`
}

func (b *XBorder) Equal(other *XBorder) bool {
	if b == nil && other == nil {
		return true
	}
	var self *XBorder
	if b == nil {
		self = &XBorder{}
	} else {
		self = b
	}
	if other == nil {
		other = &XBorder{}
	}
	return self.DiagonalDown == other.DiagonalDown &&
		self.DiagonalUp == other.DiagonalUp &&
		self.Outline == other.Outline &&
		self.Left.Equal(other.Left) &&
		self.Right.Equal(other.Right) &&
		self.Top.Equal(other.Top) &&
		self.Bottom.Equal(other.Bottom)
}

// XBorderInfo BorderInfo type
type XBorderInfo struct {
	Style string  `xml:"style,attr,omitempty"`
	Color *XColor `xml:"color,omitempty"`
}

func (b *XBorderInfo) Equal(other *XBorderInfo) bool {
	if b == nil && other == nil {
		return true
	}
	var self *XBorderInfo
	if b == nil {
		self = &XBorderInfo{}
	} else {
		self = b
	}
	if other == nil {
		other = &XBorderInfo{}
	}
	return self.Style == other.Style &&
		self.Color.Equal(other.Color)
}

// XCellStyleXfs CellStyleXfs type
type XCellStyleXfs struct {
	Count int `xml:"count,attr"`

	Xf []*XXf `xml:"xf"`
}

// XCellXfs CellFormats.
// This element contains the master formatting records (xf) which define the formatting applied to cells in this workbook. These records are the starting point for determining the formatting for a cell. Cells in the Sheet Part reference the xf records by zero-based index.
type XCellXfs struct {
	Count int    `xml:"count,attr"` // Format Count. Count of xf elements.
	Xf    []*XXf `xml:"xf"`         // Formats
}

// XXf CellFormat. Formatting Elements
// https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellformat?view=openxml-2.8.1
type XXf struct {
	NumFmtID          int         `xml:"numFmtId,attr"`                    // Number Format Id. Id of the number format (numFmt) record used by this cell format.
	FontID            int         `xml:"fontId,attr"`                      // Font Id. Zero-based index of the font record used by this cell format.
	FillID            int         `xml:"fillId,attr"`                      // Fill Id. Zero-based index of the fill record used by this cell format.
	BorderID          int         `xml:"borderId,attr"`                    // Border Id. Zero-based index of the border record used by this cell format.
	XfID              int         `xml:"xfId,attr,omitempty"`              // Format Id. For xf records contained in cellXfs this is the zero-based index of an xf record contained in cellStyleXfs corresponding to the cell style applied to the cell.
	ApplyAlignment    BoolAttr    `xml:"applyAlignment,attr,omitempty"`    // Apply Alignment. A boolean value indicating whether the alignment formatting specified for this xf should be applied.
	ApplyBorder       BoolAttr    `xml:"applyBorder,attr,omitempty"`       // Apply Border. A boolean value indicating whether the border formatting specified for this xf should be applied.
	ApplyFill         BoolAttr    `xml:"applyFill,attr,omitempty"`         // Apply Fill. A boolean value indicating whether the fill formatting specified for this xf should be applied.
	ApplyFont         BoolAttr    `xml:"applyFont,attr,omitempty"`         // Apply Font. A boolean value indicating whether the font formatting specified for this xf should be applied.
	ApplyNumberFormat BoolAttr    `xml:"applyNumberFormat,attr,omitempty"` // Apply Number Format. A boolean value indicating whether the number formatting specified for this xf should be applied.
	ApplyProtection   BoolAttr    `xml:"applyProtection,attr,omitempty"`   // Apply Protection. A boolean value indicating whether the protection formatting specified for this xf should be applied.
	PivotButton       BoolAttr    `xml:"pivotButton,attr,omitempty"`       // Pivot Button. A boolean value indicating whether the cell rendering includes a pivot table dropdown button.
	QuotePrefix       BoolAttr    `xml:"quotePrefix,attr,omitempty"`       // Quote Prefix. A boolean value indicating whether the text string in a cell should be prefixed by a single quote mark (e.g., 'text). In these cases, the quote is not stored in the Shared Strings Part.
	Alignment         *XAlignment `xml:"alignment,omitempty"`              // Alignment
}

func (f *XXf) Equal(other *XXf) bool {
	if f == nil && other == nil {
		return true
	}
	var self *XXf
	if f == nil {
		self = &XXf{}
	} else {
		self = f
	}
	if other == nil {
		other = &XXf{}
	}
	return self.NumFmtID == other.NumFmtID &&
		self.FontID == other.FontID &&
		self.FillID == other.FillID &&
		self.BorderID == other.BorderID &&
		self.XfID == other.XfID &&
		self.ApplyAlignment == other.ApplyAlignment &&
		self.ApplyBorder == other.ApplyBorder &&
		self.ApplyFill == other.ApplyFill &&
		self.ApplyFont == other.ApplyFont &&
		self.ApplyNumberFormat == other.ApplyNumberFormat &&
		self.ApplyProtection == other.ApplyProtection &&
		self.PivotButton == other.PivotButton &&
		self.QuotePrefix == other.QuotePrefix &&
		self.Alignment.Equal(other.Alignment)
}

// XAlignment Alignment type. Formatting information pertaining to text alignment in cells. There are a variety of choices for how text is aligned both horizontally and vertically, as well as indentation settings, and so on.
// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.alignment?view=openxml-2.8.1
type XAlignment struct {
	Horizontal      string   `xml:"horizontal,attr,omitempty"`      // Specifies the type of horizontal alignment in cells.
	Indent          int      `xml:"indent,attr,omitempty"`          // An integer value, where an increment of 1 represents 3 spaces. Indicates the number of spaces (of the normal style font) of indentation for text in a cell.
	JustifyLastLine BoolAttr `xml:"justifyLastLine,attr,omitempty"` // A boolean value indicating if the cells justified or distributed alignment should be used on the last line of text.
	ReadingOrder    int      `xml:"readingOrder,attr,omitempty"`    // An integer value indicating whether the reading order (bidirectionality) of the cell is left-to-right, right-to-left, or context dependent.
	RelativeIndent  int      `xml:"relativeIndent,attr,omitempty"`  // An integer value (used only in a dxf element) to indicate the additional number of spaces of indentation to adjust for text in a cell.
	ShrinkToFit     BoolAttr `xml:"shrinkToFit,attr,omitempty"`     // A boolean value indicating if the displayed text in the cell should be shrunk to fit the cell width. Not applicable when a cell contains multiple lines of text.
	TextRotation    int      `xml:"textRotation,attr,omitempty"`    // Text rotation in cells. Expressed in degrees. Values range from 0 to 180. The first letter of the text is considered the center-point of the arc.
	Vertical        string   `xml:"vertical,attr,omitempty"`        // Vertical alignment in cells.
	WrapText        BoolAttr `xml:"wrapText,attr,omitempty"`        // A boolean value indicating if the text in a cell should be line-wrapped within the cell.
}

func (a *XAlignment) Equal(other *XAlignment) bool {
	if a == nil && other == nil {
		return true
	}
	var self *XAlignment
	if a == nil {
		self = &XAlignment{}
	} else {
		self = a
	}
	if other == nil {
		other = &XAlignment{}
	}
	return self.Horizontal == other.Horizontal &&
		self.Indent == other.Indent &&
		self.JustifyLastLine == other.JustifyLastLine &&
		self.ReadingOrder == other.ReadingOrder &&
		self.RelativeIndent == other.RelativeIndent &&
		self.ShrinkToFit == other.ShrinkToFit &&
		self.TextRotation == other.TextRotation &&
		self.Vertical == other.Vertical &&
		self.WrapText == other.WrapText
}

func (a *XAlignment) IsZero() bool {
	return a == nil || a.Equal(&XAlignment{})
}

// XCellStyles
// This element contains the named cell styles, consisting of a sequence of named style records.
// A named cell style is a collection of direct or themed formatting (e.g., cell border, cell fill, and font type/size/style) grouped together into a single named style, and can be applied to a cell.
//
// Example:
//
//	For example, "Normal", "Heading 1", "Title", and "20% Accent1" are named cell styles expressed below.
//	They have builtInId's associated with them, and use xfId to reference the specific formatting elements pertaining to the particular style.
//	The xfId is a zero-based index, referencing an xf record in the cellStyleXfs collection.
//
//	<cellStyles count="4">
//	  <cellStyle name="20% - Accent1" xfId="3" builtinId="30"/>
//	  <cellStyle name="Heading 1" xfId="2" builtinId="16"/>
//	  <cellStyle name="Normal" xfId="0" builtinId="0"/>
//	  <cellStyle name="Title" xfId="1" builtinId="15"/>
//	</cellStyles>
type XCellStyles struct {
	Count int `xml:"count,attr"` // Style Count

	CellStyle []*XCellStyle `xml:"cellStyle"`
}

// XCellStyle CellStyle type
type XCellStyle struct {
	Name          string `xml:"name,attr"`                    // User Defined Cell Style
	XfID          int    `xml:"xfId,attr"`                    // Format Id
	BuiltinID     int    `xml:"builtinId,attr"`               // Built-In Style Id
	CustomBuiltin string `xml:"customBuiltin,attr,omitempty"` // Custom Built In
	Hidden        string `xml:"hidden,attr,omitempty"`        // Hidden Style
	ILevel        string `xml:"iLevel,attr,omitempty"`        // Outline Style
}

// XDxfs Dxfs type
type XDxfs struct {
	Count int `xml:"count,attr"`
}

// XTableStyles TableStyles type
type XTableStyles struct {
	Count             int    `xml:"count,attr"`
	DefaultTableStyle string `xml:"defaultTableStyle,attr"`
	DefaultPivotStyle string `xml:"defaultPivotStyle,attr"`
}

// XStyleSheetExtLst StyleSheetExtLst type
type XStyleSheetExtLst struct {
	Ext []*XStyleSheetExt `xml:"ext"`
}

// XStyleSheetExt StyleSheetExt type
type XStyleSheetExt struct {
	URI            string           `xml:"uri,attr"`
	XmlnsX14       string           `xml:"xmlns:x14,attr,omitempty"`
	XmlnsX15       string           `xml:"xmlns:x15,attr,omitempty"`
	SlicerStyles   *XSlicerStyles   `xml:"x14:slicerStyles,omitempty"`
	TimelineStyles *XTimelineStyles `xml:"x15:timelineStyles,omitempty"`
}

// XSlicerStyles SlicerStyles type
type XSlicerStyles struct {
	DefaultSlicerStyle string `xml:"defaultSlicerStyle,attr"`
}

// XTimelineStyles TimelineStyles type
type XTimelineStyles struct {
	DefaultTimelineStyle string `xml:"defaultTimelineStyle,attr"`
}

// XColors Colors type
type XColors struct {
	IndexedColors []any `xml:"x:indexedColors"` // Color Indexes
	MruColors     []any `xml:"x:mruColors"`     // MRU Colors
}

// XNumFmts This element defines the number formats in this workbook, consisting of a sequence of numFmt records, where each numFmt record defines a particular number format, indicating how to format and render the numeric value of a cell.
// https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.numberingformats?view=openxml-2.8.1
type XNumFmts struct {
	Count  int        `xml:"count,attr"` // Number Format Count. Count of number format elements.
	NumFmt []*XNumFmt `xml:"numFmt"`
}

// XNumFmt This element specifies number format properties which indicate how to format and render the numeric value of a cell.
// https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.numberingformat?view=openxml-2.8.1
type XNumFmt struct {
	FormatCode string `xml:"formatCode,attr"` // Number Format Code
	NumFmtID   int    `xml:"numFmtId,attr"`   // Number Format Id. Id used by the master style records (xf's) to reference this number format.
}

func (n *XNumFmt) Equal(other *XNumFmt) bool {
	if n == nil && other == nil {
		return true
	}
	var self *XNumFmt
	if n == nil {
		self = &XNumFmt{}
	} else {
		self = n
	}
	if other == nil {
		other = &XNumFmt{}
	}
	return self.FormatCode == other.FormatCode &&
		self.NumFmtID == other.NumFmtID
}

// XFormat PivotTable Format. Represents the format defined in the PivotTable.
// https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.format?view=openxml-2.8.1
type XFormat struct {
	// Action Format Action. Specifies the formatting behavior for the area indicated in the pivotArea element.
	// The default value for this attribute is "formatting," which indicates that the specified cells have some formatting applied.
	// The format is specified in the dxfId attribute. If the formatting is cleared from the cells, then the value of this attribute becomes "blank."
	Action string `xml:"action,attr"`
	// DxfID Format Id. Specifies the identifier of the format the application is currently using for the PivotTable.
	// Formatting information is written to the styles part. See the Styles section (§18.8) for more information on formats.
	DxfID int `xml:"dxfId,attr"`
	//ExtLst []any `xml:"extLst"` // Future Feature Data Storage Area
	//PivotArea []any `xml:"pivotArea"` // Pivot Area
}

// NewDefaultXStyleSheet create *XStyleSheet with default template
func NewDefaultXStyleSheet(defaultFontName string, defaultFontSize int) *XStyleSheet {
	return &XStyleSheet{
		XmlnsMc:     "http://schemas.openxmlformats.org/markup-compatibility/2006",
		McIgnorable: "x14ac x16r2",
		XmlnsX14ac:  "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
		XmlnsX16r2:  "http://schemas.microsoft.com/office/spreadsheetml/2015/02/main",

		Fonts: &XFonts{
			Count:      1,
			KnownFonts: "1",
			Font: []*XStyleSheetFont{
				{
					Sz:      &XIntValAttr{Val: defaultFontSize},
					Color:   &XColor{Theme: 1},
					Name:    &XValAttrElement{Val: defaultFontName},
					Family:  &XIntValAttr{Val: 0},
					Charset: &XIntValAttr{Val: 0},
					// Scheme: &XValAttrElement{Val: "minor"},
				},
			},
		},
		Fills: &XFills{
			Count: 2,
			Fill: []*XFill{
				{PatternFill: &XPatternFill{PatternType: "none"}},
				{PatternFill: &XPatternFill{PatternType: "gray125"}},
			},
		},
		Borders: &XBorders{
			Count: 1,
			Border: []*XBorder{
				{
					Left:     &XBorderInfo{},
					Right:    &XBorderInfo{},
					Top:      &XBorderInfo{},
					Bottom:   &XBorderInfo{},
					Diagonal: &XBorderInfo{},
				},
			},
		},
		CellStyleXfs: &XCellStyleXfs{
			Count: 1,
			Xf: []*XXf{
				{
					NumFmtID: 0,
					FontID:   0,
					FillID:   0,
					BorderID: 0,
				},
			},
		},
		CellXfs: &XCellXfs{
			Count: 1,
			Xf: []*XXf{
				{
					NumFmtID: 0,
					FontID:   0,
					FillID:   0,
					BorderID: 0,
					XfID:     0,
				},
			},
		},
		CellStyles: &XCellStyles{
			Count: 1,
			CellStyle: []*XCellStyle{
				{
					Name:      "Normal",
					XfID:      0,
					BuiltinID: 0,
				},
			},
		},
		Dxfs: &XDxfs{Count: 0},
		TableStyles: &XTableStyles{
			Count:             0,
			DefaultTableStyle: "TableStyleMedium2",
			DefaultPivotStyle: "PivotStyleLight16",
		},
		ExtLst: &XStyleSheetExtLst{
			Ext: []*XStyleSheetExt{
				{
					URI:          "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}",
					XmlnsX14:     "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
					SlicerStyles: &XSlicerStyles{DefaultSlicerStyle: "SlicerStyleLight1"},
				},
				{
					URI:            "{9260A510-F301-46a8-8635-F512D64BE5F5}",
					XmlnsX15:       "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main",
					TimelineStyles: &XTimelineStyles{DefaultTimelineStyle: "TimeSlicerStyleLight1"},
				},
			},
		},
	}
}
