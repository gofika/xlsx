package packaging

import "encoding/xml"

// XThemeU fix XML ns
type XThemeU struct {
	XMLName xml.Name `xml:"theme"`
	XmlnsA  string   `xml:"a,attr"`
	Name    string   `xml:"name,attr"`

	ThemeElements     *XThemeElementsU     `xml:"themeElements"`
	ObjectDefaults    *XObjectDefaultsU    `xml:"objectDefaults"`
	ExtraClrSchemeLst *XExtraClrSchemeLstU `xml:"extraClrSchemeLst"`
	ExtLst            *XExtLstU            `xml:"extLst"`
}

// XThemeElementsU fix XML ns
type XThemeElementsU struct {
	ClrScheme  *XClrSchemeU  `xml:"clrScheme"`
	FontScheme *XFontSchemeU `xml:"fontScheme"`
	FmtScheme  *XFmtSchemeU  `xml:"fmtScheme"`
}

// XClrSchemeU fix XML ns
type XClrSchemeU struct {
	Name     string            `xml:"name,attr"`
	Dk1      *XSysClrElementU  `xml:"dk1"`
	Lt1      *XSysClrElementU  `xml:"lt1"`
	Dk2      *XSrgbClrElementU `xml:"dk2"`
	Lt2      *XSrgbClrElementU `xml:"lt2"`
	Accent1  *XSrgbClrElementU `xml:"accent1"`
	Accent2  *XSrgbClrElementU `xml:"accent2"`
	Accent3  *XSrgbClrElementU `xml:"accent3"`
	Accent4  *XSrgbClrElementU `xml:"accent4"`
	Accent5  *XSrgbClrElementU `xml:"accent5"`
	Accent6  *XSrgbClrElementU `xml:"accent6"`
	Hlink    *XSrgbClrElementU `xml:"hlink"`
	FolHlink *XSrgbClrElementU `xml:"folHlink"`
}

// XSysClrElementU fix XML ns
type XSysClrElementU struct {
	SysClr *XSysClrU `xml:"sysClr"`
}

// XSysClrU fix XML ns
type XSysClrU struct {
	Val     string `xml:"val,attr"`
	LastClr string `xml:"lastClr,attr"`
}

// XSrgbClrElementU fix XML ns
type XSrgbClrElementU struct {
	SrgbClr *XSrgbClrU `xml:"srgbClr"`
}

// XSrgbClrU fix XML ns
type XSrgbClrU struct {
	Val   string            `xml:"val,attr"`
	Alpha *XValAttrElementU `xml:"alpha,omitempty"`
}

// XFontSchemeU fix XML ns
type XFontSchemeU struct {
	Name      string         `xml:"name,attr"`
	MajorFont *XFontElementU `xml:"majorFont"`
	MinorFont *XFontElementU `xml:"minorFont"`
}

// XFontElementU fix XML ns
type XFontElementU struct {
	Latin *XLatinU       `xml:"latin"`
	Ea    *XTypefaceU    `xml:"ea"`
	Cs    *XTypefaceU    `xml:"cs"`
	Font  []*XThemeFontU `xml:"font"`
}

// XLatinU fix XML ns
type XLatinU struct {
	Typeface string `xml:"typeface,attr"`
	Panose   string `xml:"panose,attr"`
}

// XTypefaceU fix XML ns
type XTypefaceU struct {
	Typeface string `xml:"typeface,attr"`
}

// XThemeFontU fix XML ns
type XThemeFontU struct {
	Script   string `xml:"script,attr"`
	Typeface string `xml:"typeface,attr"`
}

// XFmtSchemeU fix XML ns
type XFmtSchemeU struct {
	Name           string            `xml:"name,attr"`
	FillStyleLst   *XFillStyleLstU   `xml:"fillStyleLst"`
	LnStyleLst     *XLnStyleLstU     `xml:"lnStyleLst"`
	EffectStyleLst *XEffectStyleLstU `xml:"effectStyleLst"`
	BgFillStyleLst *XBgFillStyleLstU `xml:"bgFillStyleLst"`
}

// XFillStyleLstU fix XML ns
type XFillStyleLstU struct {
	SolidFill *XSolidFillU  `xml:"solidFill"`
	GradFill  []*XGradFillU `xml:"gradFill"`
}

// XSolidFillU fix XML ns
type XSolidFillU struct {
	SchemeClr *XSchemeClrU      `xml:"schemeClr"`
	Tint      *XValAttrElementU `xml:"tint,omitempty"`
	SatMod    *XValAttrElementU `xml:"satMod,omitempty"`
}

// XSchemeClrU fix XML ns
type XSchemeClrU struct {
	Val    string            `xml:"val,attr"`
	LumMod *XValAttrElementU `xml:"lumMod,omitempty"`
	SatMod *XValAttrElementU `xml:"satMod,omitempty"`
	Tint   *XValAttrElementU `xml:"tint,omitempty"`
	Shade  *XValAttrElementU `xml:"shade,omitempty"`
}

// XGradFillU fix XML ns
type XGradFillU struct {
	RotWithShape string   `xml:"rotWithShape,attr"`
	GsLst        *XGsLstU `xml:"gsLst"`
	Lin          *XLinU   `xml:"lin"`
}

// XGsLstU fix XML ns
type XGsLstU struct {
	Gs []*XGsU `xml:"gs"`
}

// XGsU fix XML ns
type XGsU struct {
	Pos       string       `xml:"pos,attr"`
	SchemeClr *XSchemeClrU `xml:"schemeClr"`
}

// XValAttrElementU fix XML ns
type XValAttrElementU struct {
	Val string `xml:"val,attr"`
}

// XLinU fix XML ns
type XLinU struct {
	Ang    string `xml:"ang,attr"`
	Scaled string `xml:"scaled,attr"`
}

// XLnStyleLstU fix XML ns
type XLnStyleLstU struct {
	Ln []*XLnU `xml:"ln"`
}

// XLnU fix XML ns
type XLnU struct {
	W         string            `xml:"w,attr"`
	Cap       string            `xml:"cap,attr"`
	Cmpd      string            `xml:"cmpd,attr"`
	Algn      string            `xml:"algn,attr"`
	SolidFill *XSolidFillU      `xml:"solidFill"`
	PrstDash  *XValAttrElementU `xml:"prstDash"`
	Miter     *XMiterU          `xml:"miter"`
}

// XMiterU fix XML ns
type XMiterU struct {
	Lim string `xml:"lim,attr"`
}

// XEffectStyleLstU fix XML ns
type XEffectStyleLstU struct {
	EffectStyle []*XEffectStyleU `xml:"effectStyle"`
}

// XEffectStyleU fix XML ns
type XEffectStyleU struct {
	EffectLst *XEffectLstU `xml:"effectLst"`
}

// XEffectLstU fix XML ns
type XEffectLstU struct {
	OuterShdw *XOuterShdwU `xml:"outerShdw"`
}

// XOuterShdwU fix XML ns
type XOuterShdwU struct {
	BlurRad      string     `xml:"blurRad,attr"`
	Dist         string     `xml:"dist,attr"`
	Dir          string     `xml:"dir,attr"`
	Algn         string     `xml:"algn,attr"`
	RotWithShape string     `xml:"rotWithShape,attr"`
	SrgbClr      *XSrgbClrU `xml:"srgbClr"`
}

// XBgFillStyleLstU fix XML ns
type XBgFillStyleLstU struct {
	SolidFill []*XSolidFillU `xml:"solidFill"`
	GradFill  *XGradFillU    `xml:"gradFill"`
}

// XObjectDefaultsU fix XML ns
type XObjectDefaultsU struct {
}

// XExtraClrSchemeLstU fix XML ns
type XExtraClrSchemeLstU struct {
}

// XExtLstU fix XML ns
type XExtLstU struct {
	Ext *XExtU `xml:"ext"`
}

// XExtU fix XML ns
type XExtU struct {
	URI         string         `xml:"uri,attr"`
	ThemeFamily *XThemeFamilyU `xml:"themeFamily"`
}

// XThemeFamilyU fix XML ns
type XThemeFamilyU struct {
	XmlnsThm15 string `xml:"thm15,attr"`
	Name       string `xml:"name,attr"`
	ID         string `xml:"id,attr"`
	Vid        string `xml:"vid,attr"`
}
