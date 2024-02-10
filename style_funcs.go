package xlsx

import (
	"github.com/gofika/xlsx/packaging"
	"github.com/shopspring/decimal"
)

func colorFromPackaing(color *packaging.XColor) Color {
	if color == nil {
		return Color{}
	}
	tint, err := decimal.NewFromString(color.Tint)
	if err != nil {
		tint = decimal.Zero
	}
	return Color{
		ThemeColor: ThemeColor(color.Theme),
		Indexed:    int(color.Indexed),
		Tint:       tint,
		Color:      color.RGB,
	}
}

func packaingToColor(color Color) *packaging.XColor {
	c := &packaging.XColor{}
	if color.ThemeColor != 0 {
		c.Theme = packaging.OmitIntAttr(color.ThemeColor)
	}
	if color.Indexed != 0 {
		c.Indexed = packaging.OmitIntAttr(color.Indexed)
	}
	if !color.Tint.IsZero() {
		c.Tint = color.Tint.String()
	}
	if color.Color != "" {
		c.RGB = color.Color
	}
	return c
	// return &packaging.XColor{
	// 	Theme:   int(color.ThemeColor),
	// 	Indexed: color.Indexed,
	// 	Tint:    color.Tint,
	// 	RGB:     color.Color,
	// }
}

func fontFromPackaing(font *packaging.XStyleSheetFont) Font {
	if font == nil {
		return Font{}
	}
	return Font{
		Bold:                font.B.Value(),
		FontCharSet:         FontCharSet(font.Charset.Value()),
		FontColor:           colorFromPackaing(font.Color),
		Condense:            font.Condense.Value(),
		Extend:              font.Extend.Value(),
		FontFamilyNumbering: FontFamilyNumbering(font.Family.Value()),
		Italic:              font.I.Value(),
		FontName:            font.Name.Value(),
		Outline:             font.Outline.Value(),
		FontScheme:          FontScheme(font.Scheme.Value()),
		Shadow:              font.Shadow.Value(),
		Strikethrough:       font.Strike.Value(),
		FontSize:            font.Sz.Value(),
		Underline:           FontUnderline(font.U.Value()),
		VerticalAlignment:   FontVerticalTextAlignment(font.VertAlign.Value()),
	}
}

func packaingToFont(font Font) *packaging.XStyleSheetFont {
	f := &packaging.XStyleSheetFont{}
	if font.Bold {
		f.B = &packaging.XBoolValAttr{Val: packaging.NewBool(font.Bold)}
	}
	if font.FontCharSet != 0 {
		f.Charset = &packaging.XIntValAttr{Val: int(font.FontCharSet)}
	}
	if !font.FontColor.IsZero() {
		f.Color = packaingToColor(font.FontColor)
	}
	if font.Condense {
		f.Condense = &packaging.XBoolValAttr{Val: packaging.NewBool(font.Condense)}
	}
	if font.Extend {
		f.Extend = &packaging.XBoolValAttr{Val: packaging.NewBool(font.Extend)}
	}
	if font.FontFamilyNumbering != 0 {
		f.Family = &packaging.XIntValAttr{Val: int(font.FontFamilyNumbering)}
	}
	if font.Italic {
		f.I = &packaging.XBoolValAttr{Val: packaging.NewBool(font.Italic)}
	}
	if font.FontName != "" {
		f.Name = &packaging.XValAttrElement{Val: font.FontName}
	}
	if font.Outline {
		f.Outline = &packaging.XBoolValAttr{Val: packaging.NewBool(font.Outline)}
	}
	if font.FontScheme != "" {
		f.Scheme = &packaging.XValAttrElement{Val: string(font.FontScheme)}
	}
	if font.Shadow {
		f.Shadow = &packaging.XBoolValAttr{Val: packaging.NewBool(font.Shadow)}
	}
	if font.Strikethrough {
		f.Strike = &packaging.XBoolValAttr{Val: packaging.NewBool(font.Strikethrough)}
	}
	if font.FontSize != 0 {
		f.Sz = &packaging.XIntValAttr{Val: font.FontSize}
	}
	if font.Underline != "" {
		f.U = &packaging.XValAttrElement{Val: string(font.Underline)}
	}
	if font.VerticalAlignment != "" {
		f.VertAlign = &packaging.XValAttrElement{Val: string(font.VerticalAlignment)}
	}
	return f
	// return &packaging.XStyleSheetFont{
	// 	B:         &packaging.XBoolValAttr{Val: packaging.NewBool(font.Bold)},
	// 	Charset:   &packaging.XIntValAttr{Val: int(font.FontCharSet)},
	// 	Color:     packaingToColor(font.FontColor),
	// 	Condense:  &packaging.XBoolValAttr{Val: packaging.NewBool(font.Condense)},
	// 	Extend:    &packaging.XBoolValAttr{Val: packaging.NewBool(font.Extend)},
	// 	Family:    &packaging.XIntValAttr{Val: int(font.FontFamilyNumbering)},
	// 	I:         &packaging.XBoolValAttr{Val: packaging.NewBool(font.Italic)},
	// 	Name:      &packaging.XValAttrElement{Val: font.FontName},
	// 	Outline:   &packaging.XBoolValAttr{Val: packaging.NewBool(font.Outline)},
	// 	Scheme:    &packaging.XValAttrElement{Val: string(font.FontScheme)},
	// 	Shadow:    &packaging.XBoolValAttr{Val: packaging.NewBool(font.Shadow)},
	// 	Strike:    &packaging.XBoolValAttr{Val: packaging.NewBool(font.Strikethrough)},
	// 	Sz:        &packaging.XIntValAttr{Val: font.FontSize},
	// 	U:         &packaging.XValAttrElement{Val: string(font.Underline)},
	// 	VertAlign: &packaging.XValAttrElement{Val: string(font.VerticalAlignment)},
	// }
}

func alignmentFromPackaging(alignment *packaging.XAlignment) Alignment {
	if alignment == nil {
		return Alignment{}
	}
	return Alignment{
		Horizontal:      HorizontalAlignment(alignment.Horizontal),
		Indent:          alignment.Indent,
		JustifyLastLine: alignment.JustifyLastLine.Value(),
		ReadingOrder:    AlignmentReadingOrder(alignment.ReadingOrder),
		RelativeIndent:  alignment.RelativeIndent,
		ShrinkToFit:     alignment.ShrinkToFit.Value(),
		TextRotation:    alignment.TextRotation,
		Vertical:        VerticalAlignment(alignment.Vertical),
		WrapText:        alignment.WrapText.Value(),
	}
}

func packaingToAlignment(alignment Alignment) *packaging.XAlignment {
	return &packaging.XAlignment{
		Horizontal:      string(alignment.Horizontal),
		Indent:          alignment.Indent,
		JustifyLastLine: packaging.NewBool(alignment.JustifyLastLine),
		ReadingOrder:    int(alignment.ReadingOrder),
		RelativeIndent:  alignment.RelativeIndent,
		ShrinkToFit:     packaging.NewBool(alignment.ShrinkToFit),
		TextRotation:    alignment.TextRotation,
		Vertical:        string(alignment.Vertical),
		WrapText:        packaging.NewBool(alignment.WrapText),
	}
}

func borderFromPackaging(border *packaging.XBorder) Border {
	var b Border
	if border == nil {
		return Border{}
	}
	b.DiagonalDown = border.DiagonalDown.Value()
	b.DiagonalUp = border.DiagonalUp.Value()
	b.Outline = border.Outline.Value()
	if border.Left != nil {
		b.LeftBorderColor = colorFromPackaing(border.Left.Color)
		b.LeftBorder = BorderStyle(border.Left.Style)
	}
	if border.Right != nil {
		b.RightBorderColor = colorFromPackaing(border.Right.Color)
		b.RightBorder = BorderStyle(border.Right.Style)
	}
	if border.Top != nil {
		b.TopBorderColor = colorFromPackaing(border.Top.Color)
		b.TopBorder = BorderStyle(border.Top.Style)
	}
	if border.Bottom != nil {
		b.BottomBorderColor = colorFromPackaing(border.Bottom.Color)
		b.BottomBorder = BorderStyle(border.Bottom.Style)
	}
	if border.Diagonal != nil {
		b.DiagonalBorderColor = colorFromPackaing(border.Diagonal.Color)
		b.DiagonalBorder = BorderStyle(border.Diagonal.Style)
	}
	return b
}

func packaingToBorder(border Border) *packaging.XBorder {
	var b packaging.XBorder
	if border.LeftBorder != "" {
		b.Left = &packaging.XBorderInfo{
			Color: packaingToColor(border.LeftBorderColor),
			Style: string(border.LeftBorder),
		}
	}
	if border.RightBorder != "" {
		b.Right = &packaging.XBorderInfo{
			Color: packaingToColor(border.RightBorderColor),
			Style: string(border.RightBorder),
		}
	}
	if border.TopBorder != "" {
		b.Top = &packaging.XBorderInfo{
			Color: packaingToColor(border.TopBorderColor),
			Style: string(border.TopBorder),
		}
	}
	if border.BottomBorder != "" {
		b.Bottom = &packaging.XBorderInfo{
			Color: packaingToColor(border.BottomBorderColor),
			Style: string(border.BottomBorder),
		}
	}
	if border.DiagonalBorder != "" {
		b.Diagonal = &packaging.XBorderInfo{
			Color: packaingToColor(border.DiagonalBorderColor),
			Style: string(border.DiagonalBorder),
		}
	}
	b.DiagonalDown = packaging.NewBool(border.DiagonalDown)
	b.DiagonalUp = packaging.NewBool(border.DiagonalUp)
	b.Outline = packaging.NewBool(border.Outline)
	return &b
}

func fillFromPackaging(fill *packaging.XFill) Fill {
	if fill == nil {
		return Fill{}
	}
	return Fill{
		PatternType:     FillPattern(fill.GetPatternType()),
		BackgroundColor: colorFromPackaing(fill.PatternFill.BackgroundColor),
		PatternColor:    colorFromPackaing(fill.PatternFill.ForegroundColor),
	}
}

func packaingToFill(fill Fill) *packaging.XFill {
	return &packaging.XFill{
		PatternFill: &packaging.XPatternFill{
			PatternType:     string(fill.PatternType),
			BackgroundColor: packaingToColor(fill.BackgroundColor),
			ForegroundColor: packaingToColor(fill.PatternColor),
		},
	}
}

func numberFormatFromPackaging(numberFormat *packaging.XNumFmt) NumberFormat {
	if numberFormat == nil {
		return NumberFormat{
			NumberFormatID: 0,
			Format:         builtInNumFmt[0],
		}
	}
	return NumberFormat{
		NumberFormatID: numberFormat.NumFmtID,
		Format:         numberFormat.FormatCode,
	}
}

func packaingToNumberFormat(nf NumberFormat) *packaging.XNumFmt {
	return &packaging.XNumFmt{
		NumFmtID:   nf.NumberFormatID,
		FormatCode: nf.Format,
	}
}
