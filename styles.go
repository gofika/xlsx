package xlsx

import (
	"github.com/shopspring/decimal"
)

const BuiltInNumFmtMax = 163 // The maximum number of built-in number formats.

var builtInNumFmt = map[int]string{
	0:  "", // general
	1:  "0",
	2:  "0.00",
	3:  "#,##0",
	4:  "#,##0.00",
	9:  "0%",
	10: "0.00%",
	11: "0.00e+00",
	12: "# ?/?",
	13: "# ??/??",
	14: "mm-dd-yy",
	15: "d-mmm-yy",
	16: "d-mmm",
	17: "mmm-yy",
	18: "h:mm am/pm",
	19: "h:mm:ss am/pm",
	20: "h:mm",
	21: "h:mm:ss",
	22: "m/d/yy h:mm",
	37: "#,##0 ;(#,##0)",
	38: "#,##0 ;[red](#,##0)",
	39: "#,##0.00;(#,##0.00)",
	40: "#,##0.00;[red](#,##0.00)",
	41: `_(* #,##0_);_(* \(#,##0\);_(* "-"_);_(@_)`,
	42: `_("$"* #,##0_);_("$* \(#,##0\);_("$"* "-"_);_(@_)`,
	43: `_(* #,##0.00_);_(* \(#,##0.00\);_(* "-"??_);_(@_)`,
	44: `_("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)`,
	45: "mm:ss",
	46: "[h]:mm:ss",
	47: "mmss.0",
	48: "##0.0e+0",
	49: "@",
}

type HorizontalAlignment string

const (
	HorizontalAlignmentGeneral          HorizontalAlignment = "general"
	HorizontalAlignmentCenter           HorizontalAlignment = "center"
	HorizontalAlignmentLeft             HorizontalAlignment = "left"
	HorizontalAlignmentRight            HorizontalAlignment = "right"
	HorizontalAlignmentFill             HorizontalAlignment = "fill"
	HorizontalAlignmentJustify          HorizontalAlignment = "justify"
	HorizontalAlignmentCenterContinuous HorizontalAlignment = "centerContinuous"
	HorizontalAlignmentDistributed      HorizontalAlignment = "distributed"
)

type VerticalAlignment string

const (
	VerticalAlignmentGeneral     VerticalAlignment = "bottom"
	VerticalAlignmentTop         VerticalAlignment = "top"
	VerticalAlignmentBottom      VerticalAlignment = "bottom"
	VerticalAlignmentCenter      VerticalAlignment = "center"
	VerticalAlignmentJustify     VerticalAlignment = "justify"
	VerticalAlignmentDistributed VerticalAlignment = "distributed"
)

type AlignmentReadingOrder int

const (
	AlignmentReadingOrderContextDependent AlignmentReadingOrder = iota
	AlignmentReadingOrderLeftToRight
	AlignmentReadingOrderRightToLeft
)

type Alignment struct {
	Horizontal      HorizontalAlignment   // Specifies the type of horizontal alignment in cells.
	Indent          int                   // An integer value, where an increment of 1 represents 3 spaces. Indicates the number of spaces (of the normal style font) of indentation for text in a cell.
	JustifyLastLine bool                  // A boolean value indicating if the cells justified or distributed alignment should be used on the last line of text.
	ReadingOrder    AlignmentReadingOrder // An integer value indicating whether the reading order (bidirectionality) of the cell is left-to-right, right-to-left, or context dependent.
	RelativeIndent  int                   // An integer value (used only in a dxf element) to indicate the additional number of spaces of indentation to adjust for text in a cell.
	ShrinkToFit     bool                  // A boolean value indicating if the displayed text in the cell should be shrunk to fit the cell width. Not applicable when a cell contains multiple lines of text.
	TextRotation    int                   // Text rotation in cells. Expressed in degrees. Values range from 0 to 180. The first letter of the text is considered the center-point of the arc.
	Vertical        VerticalAlignment     // Vertical alignment in cells.
	WrapText        bool                  // A boolean value indicating if the text in a cell should be line-wrapped within the cell.
}

type FontUnderline string

const (
	FontUnderlineDouble           FontUnderline = "double"
	FontUnderlineDoubleAccounting FontUnderline = "doubleAccounting"
	FontUnderlineNone             FontUnderline = "none"
	FontUnderlineSingle           FontUnderline = "single"
	FontUnderlineSingleAccounting FontUnderline = "singleAccounting"
)

type FontVerticalTextAlignment string

const (
	FontVerticalTextAlignmentBaseline    FontVerticalTextAlignment = "baseline"
	FontVerticalTextAlignmentSubscript   FontVerticalTextAlignment = "subscript"
	FontVerticalTextAlignmentSuperscript FontVerticalTextAlignment = "superscript"
)

// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.colorscheme?view=openxml-2.8.1
type ThemeColor int

const (
	ThemeColorDark1 ThemeColor = iota
	ThemeColorLight1
	ThemeColorDark2
	ThemeColorLight2
	ThemeColorAccent1
	ThemeColorAccent2
	ThemeColorAccent3
	ThemeColorAccent4
	ThemeColorAccent5
	ThemeColorAccent6
	ThemeColorHyperlink
	ThemeColorFollowedHyperlink
)

type Color struct {
	ThemeColor ThemeColor
	Indexed    int
	Tint       decimal.Decimal
	Color      string
}

func (c Color) IsZero() bool {
	return c.ThemeColor == 0 && c.Indexed == 0 && c.Tint.IsZero() && c.Color == ""
}

type FontFamilyNumbering int

const (
	FontFamilyNumberingNotApplicable FontFamilyNumbering = iota
	FontFamilyNumberingRoman
	FontFamilyNumberingSwiss
	FontFamilyNumberingModern
	FontFamilyNumberingScript
	FontFamilyNumberingDecorative
)

type FontCharSet int

const (
	FontCharSetAnsi        FontCharSet = 0   // ASCII character set.
	FontCharSetDefault     FontCharSet = 1   // System default character set.
	FontCharSetSymbol      FontCharSet = 2   // Symbol character set.
	FontCharSetMac         FontCharSet = 77  // Characters used by Macintosh.
	FontCharSetShiftJIS    FontCharSet = 128 // Japanese character set.
	FontCharSetHangul      FontCharSet = 129 // Korean character set.
	FontCharSetHangeul     FontCharSet = 129 // Another common spelling of the Korean character set.
	FontCharSetJohab       FontCharSet = 130 // Korean character set.
	FontCharSetGB2312      FontCharSet = 134 // Chinese character set used in mainland China.
	FontCharSetChineseBig5 FontCharSet = 136 // Chinese character set used mostly in Hong Kong SAR and Taiwan.
	FontCharSetGreek       FontCharSet = 161 // Greek character set.
	FontCharSetTurkish     FontCharSet = 162 // Turkish character set.
	FontCharSetVietnamese  FontCharSet = 163 // Vietnamese character set.
	FontCharSetHebrew      FontCharSet = 177 // Hebrew character set.
	FontCharSetArabic      FontCharSet = 178 // Arabic character set.
	FontCharSetBaltic      FontCharSet = 186 // Baltic character set.
	FontCharSetRussian     FontCharSet = 204 // Russian character set.
	FontCharSetThai        FontCharSet = 222 // Thai character set.
	FontCharSetEastEurope  FontCharSet = 238 // Eastern European character set.
	FontCharSetOEM         FontCharSet = 255 // Extended ASCII character set used with disk operating system (DOS) and some Microsoft Windows fonts.
)

type FontScheme string

const (
	FontSchemeNone  FontScheme = "none" // Not a part of theme scheme.
	FontSchemeMajor                     // A major font of a theme, generally used for headings.
	FontSchemeMinor                     // A minor font of a theme, generally used to body and paragraphs.
)

type Font struct {
	Bold                bool
	Italic              bool
	Underline           FontUnderline
	Strikethrough       bool
	VerticalAlignment   FontVerticalTextAlignment
	Shadow              bool
	FontSize            int
	FontColor           Color
	FontName            string
	FontFamilyNumbering FontFamilyNumbering
	FontCharSet         FontCharSet
	FontScheme          FontScheme
	Condense            bool
	Extend              bool
	Outline             bool
}

type BorderStyle string

const (
	BorderStyleNone             BorderStyle = "none"
	BorderStyleThin             BorderStyle = "thin"
	BorderStyleMedium           BorderStyle = "medium"
	BorderStyleDashed           BorderStyle = "dashed"
	BorderStyleDotted           BorderStyle = "dotted"
	BorderStyleThick            BorderStyle = "thick"
	BorderStyleDouble           BorderStyle = "double"
	BorderStyleHair             BorderStyle = "hair"
	BorderStyleMediumDashed     BorderStyle = "mediumDashed"
	BorderStyleDashDot          BorderStyle = "dashDot"
	BorderStyleMediumDashDot    BorderStyle = "mediumDashDot"
	BorderStyleDashDotDot       BorderStyle = "dashDotDot"
	BorderStyleMediumDashDotDot BorderStyle = "mediumDashDotDot"
	BorderStyleSlantDashDot     BorderStyle = "slantDashDot"
)

type Border struct {
	OutsideBorder       BorderStyle
	OutsideBorderColor  Color
	InsideBorder        BorderStyle
	InsideBorderColor   Color
	LeftBorder          BorderStyle
	LeftBorderColor     Color
	RightBorder         BorderStyle
	RightBorderColor    Color
	TopBorder           BorderStyle
	TopBorderColor      Color
	BottomBorder        BorderStyle
	BottomBorderColor   Color
	DiagonalBorder      BorderStyle
	DiagonalBorderColor Color
	DiagonalUp          bool
	DiagonalDown        bool
	Outline             bool
}

type FillPattern string

const (
	FillPatternNone            FillPattern = "none"
	FillPatternSolid           FillPattern = "solid"
	FillPatternMediumGray      FillPattern = "mediumGray"
	FillPatternDarkGray        FillPattern = "darkGray"
	FillPatternLightGray       FillPattern = "lightGray"
	FillPatternDarkHorizontal  FillPattern = "darkHorizontal"
	FillPatternDarkVertical    FillPattern = "darkVertical"
	FillPatternDarkDown        FillPattern = "darkDown"
	FillPatternDarkUp          FillPattern = "darkUp"
	FillPatternDarkGrid        FillPattern = "darkGrid"
	FillPatternDarkTrellis     FillPattern = "darkTrellis"
	FillPatternLightHorizontal FillPattern = "lightHorizontal"
	FillPatternLightVertical   FillPattern = "lightVertical"
	FillPatternLightDown       FillPattern = "lightDown"
	FillPatternLightUp         FillPattern = "lightUp"
	FillPatternLightGrid       FillPattern = "lightGrid"
	FillPatternLightTrellis    FillPattern = "lightTrellis"
	FillPatternGray125         FillPattern = "gray125"
	FillPatternGray0625        FillPattern = "gray0625"
)

type Fill struct {
	BackgroundColor Color
	PatternColor    Color
	PatternType     FillPattern
}

type NumberFormat struct {
	NumberFormatID int
	Format         string
}

type Style struct {
	Font               Font
	Alignment          Alignment
	Border             Border
	Fill               Fill
	IncludeQuotePrefix bool
	NumberFormat       NumberFormat
}
