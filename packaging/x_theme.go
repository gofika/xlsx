package packaging

import "encoding/xml"

// Theme Defines
const (
	ThemeContentType      = "application/vnd.openxmlformats-officedocument.theme+xml"
	ThemeRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"

	ThemePath     = "xl"
	ThemeFileName = "theme/theme%d.xml"
)

// XTheme Theme XML doc
type XTheme struct {
	XMLName xml.Name `xml:"a:theme"`
	XmlnsA  string   `xml:"xmlns:a,attr"`
	Name    string   `xml:"name,attr"`

	ThemeElements     *XThemeElements     `xml:"a:themeElements"`
	ObjectDefaults    *XObjectDefaults    `xml:"a:objectDefaults"`
	ExtraClrSchemeLst *XExtraClrSchemeLst `xml:"a:extraClrSchemeLst"`
	ExtLst            *XExtLst            `xml:"a:extLst"`
}

// XThemeElements ThemeElements type
type XThemeElements struct {
	ClrScheme  *XClrScheme  `xml:"a:clrScheme"`
	FontScheme *XFontScheme `xml:"a:fontScheme"`
	FmtScheme  *XFmtScheme  `xml:"a:fmtScheme"`
}

// XClrScheme ClrScheme type
type XClrScheme struct {
	Name     string           `xml:"name,attr"`
	Dk1      *XSysClrElement  `xml:"a:dk1"`
	Lt1      *XSysClrElement  `xml:"a:lt1"`
	Dk2      *XSrgbClrElement `xml:"a:dk2"`
	Lt2      *XSrgbClrElement `xml:"a:lt2"`
	Accent1  *XSrgbClrElement `xml:"a:accent1"`
	Accent2  *XSrgbClrElement `xml:"a:accent2"`
	Accent3  *XSrgbClrElement `xml:"a:accent3"`
	Accent4  *XSrgbClrElement `xml:"a:accent4"`
	Accent5  *XSrgbClrElement `xml:"a:accent5"`
	Accent6  *XSrgbClrElement `xml:"a:accent6"`
	Hlink    *XSrgbClrElement `xml:"a:hlink"`
	FolHlink *XSrgbClrElement `xml:"a:folHlink"`
}

// XSysClrElement SysClrElement type
type XSysClrElement struct {
	SysClr *XSysClr `xml:"a:sysClr"`
}

// XSysClr SysClr type
type XSysClr struct {
	Val     string `xml:"val,attr"`
	LastClr string `xml:"lastClr,attr"`
}

// XSrgbClrElement SrgbClrElement type
type XSrgbClrElement struct {
	SrgbClr *XSrgbClr `xml:"a:srgbClr"`
}

// XSrgbClr SrgbClr type
type XSrgbClr struct {
	Val   string           `xml:"val,attr"`
	Alpha *XValAttrElement `xml:"a:alpha,omitempty"`
}

// XFontScheme FontScheme type
type XFontScheme struct {
	Name      string        `xml:"name,attr"`
	MajorFont *XFontElement `xml:"a:majorFont"`
	MinorFont *XFontElement `xml:"a:minorFont"`
}

// XFontElement FontElement type
type XFontElement struct {
	Latin *XLatin       `xml:"a:latin"`
	Ea    *XTypeface    `xml:"a:ea"`
	Cs    *XTypeface    `xml:"a:cs"`
	Font  []*XThemeFont `xml:"a:font"`
}

// XLatin Latin type
type XLatin struct {
	Typeface string `xml:"typeface,attr"`
	Panose   string `xml:"panose,attr"`
}

// XTypeface Typeface type
type XTypeface struct {
	Typeface string `xml:"typeface,attr"`
}

// XThemeFont ThemeFont type
type XThemeFont struct {
	Script   string `xml:"script,attr"`
	Typeface string `xml:"typeface,attr"`
}

// XFmtScheme FmtScheme type
type XFmtScheme struct {
	Name           string           `xml:"name,attr"`
	FillStyleLst   *XFillStyleLst   `xml:"a:fillStyleLst"`
	LnStyleLst     *XLnStyleLst     `xml:"a:lnStyleLst"`
	EffectStyleLst *XEffectStyleLst `xml:"a:effectStyleLst"`
	BgFillStyleLst *XBgFillStyleLst `xml:"a:bgFillStyleLst"`
}

// XFillStyleLst FillStyleLst type
type XFillStyleLst struct {
	SolidFill *XSolidFill  `xml:"a:solidFill"`
	GradFill  []*XGradFill `xml:"a:gradFill"`
}

// XSolidFill SolidFill type
type XSolidFill struct {
	SchemeClr *XSchemeClr      `xml:"a:schemeClr"`
	Tint      *XValAttrElement `xml:"a:tint,omitempty"`
	SatMod    *XValAttrElement `xml:"a:satMod,omitempty"`
}

// XSchemeClr SchemeClr type
type XSchemeClr struct {
	Val    string           `xml:"val,attr"`
	LumMod *XValAttrElement `xml:"a:lumMod,omitempty"`
	SatMod *XValAttrElement `xml:"a:satMod,omitempty"`
	Tint   *XValAttrElement `xml:"a:tint,omitempty"`
	Shade  *XValAttrElement `xml:"a:shade,omitempty"`
}

// XGradFill GradFill type
type XGradFill struct {
	RotWithShape string  `xml:"rotWithShape,attr"`
	GsLst        *XGsLst `xml:"a:gsLst"`
	Lin          *XLin   `xml:"a:lin"`
}

// XGsLst GsLst type
type XGsLst struct {
	Gs []*XGs `xml:"a:gs"`
}

// XGs Gs type
type XGs struct {
	Pos       string      `xml:"pos,attr"`
	SchemeClr *XSchemeClr `xml:"a:schemeClr"`
}

// XLin Lin type
type XLin struct {
	Ang    string `xml:"ang,attr"`
	Scaled string `xml:"scaled,attr"`
}

// XLnStyleLst LnStyleLst type
type XLnStyleLst struct {
	Ln []*XLn `xml:"a:ln"`
}

// XLn Ln type
type XLn struct {
	W         string           `xml:"w,attr"`
	Cap       string           `xml:"cap,attr"`
	Cmpd      string           `xml:"cmpd,attr"`
	Algn      string           `xml:"algn,attr"`
	SolidFill *XSolidFill      `xml:"a:solidFill"`
	PrstDash  *XValAttrElement `xml:"a:prstDash"`
	Miter     *XMiter          `xml:"a:miter"`
}

// XMiter Miter type
type XMiter struct {
	Lim string `xml:"lim,attr"`
}

// XEffectStyleLst EffectStyleLst type
type XEffectStyleLst struct {
	EffectStyle []*XEffectStyle `xml:"a:effectStyle"`
}

// XEffectStyle EffectStyle type
type XEffectStyle struct {
	EffectLst *XEffectLst `xml:"a:effectLst"`
}

// XEffectLst EffectLst type
type XEffectLst struct {
	OuterShdw *XOuterShdw `xml:"a:outerShdw"`
}

// XOuterShdw OuterShdw type
type XOuterShdw struct {
	BlurRad      string    `xml:"blurRad,attr"`
	Dist         string    `xml:"dist,attr"`
	Dir          string    `xml:"dir,attr"`
	Algn         string    `xml:"algn,attr"`
	RotWithShape string    `xml:"rotWithShape,attr"`
	SrgbClr      *XSrgbClr `xml:"a:srgbClr"`
}

// XBgFillStyleLst BgFillStyleLst type
type XBgFillStyleLst struct {
	SolidFill []*XSolidFill `xml:"a:solidFill"`
	GradFill  *XGradFill    `xml:"a:gradFill"`
}

// XObjectDefaults ObjectDefaults type
type XObjectDefaults struct {
}

// XExtraClrSchemeLst ExtraClrSchemeLst type
type XExtraClrSchemeLst struct {
}

// XExtLst ExtLst type
type XExtLst struct {
	Ext *XExt `xml:"a:ext"`
}

// XExt Ext type
type XExt struct {
	URI         string        `xml:"uri,attr"`
	ThemeFamily *XThemeFamily `xml:"thm15:themeFamily"`
}

// XThemeFamily ThemeFamily type
type XThemeFamily struct {
	XmlnsThm15 string `xml:"xmlns:thm15,attr"`
	Name       string `xml:"name,attr"`
	ID         string `xml:"id,attr"`
	Vid        string `xml:"vid,attr"`
}

// NewDefaultXTheme create *XTheme with default template
func NewDefaultXTheme() *XTheme {
	return &XTheme{
		XmlnsA: "http://schemas.openxmlformats.org/drawingml/2006/main",
		Name:   "Office Theme",

		ThemeElements: &XThemeElements{
			ClrScheme: &XClrScheme{
				Name: "Office",

				Dk1: &XSysClrElement{
					SysClr: &XSysClr{
						Val:     "windowText",
						LastClr: "000000",
					},
				},
				Lt1: &XSysClrElement{
					SysClr: &XSysClr{
						Val:     "window",
						LastClr: "FFFFFF",
					},
				},
				Dk2: &XSrgbClrElement{
					SrgbClr: &XSrgbClr{
						Val: "44546A",
					},
				},
				Lt2: &XSrgbClrElement{
					SrgbClr: &XSrgbClr{
						Val: "E7E6E6",
					},
				},
				Accent1: &XSrgbClrElement{
					SrgbClr: &XSrgbClr{
						Val: "5B9BD5",
					},
				},
				Accent2: &XSrgbClrElement{
					SrgbClr: &XSrgbClr{
						Val: "ED7D31",
					},
				},
				Accent3: &XSrgbClrElement{
					SrgbClr: &XSrgbClr{
						Val: "A5A5A5",
					},
				},
				Accent4: &XSrgbClrElement{
					SrgbClr: &XSrgbClr{
						Val: "FFC000",
					},
				},
				Accent5: &XSrgbClrElement{
					SrgbClr: &XSrgbClr{
						Val: "4472C4",
					},
				},
				Accent6: &XSrgbClrElement{
					SrgbClr: &XSrgbClr{
						Val: "70AD47",
					},
				},
				Hlink: &XSrgbClrElement{
					SrgbClr: &XSrgbClr{
						Val: "0563C1",
					},
				},
				FolHlink: &XSrgbClrElement{
					SrgbClr: &XSrgbClr{
						Val: "954F72",
					},
				},
			},
			FontScheme: &XFontScheme{
				Name: "Office",
				MajorFont: &XFontElement{
					Latin: &XLatin{Typeface: "Calibri Light", Panose: "020F0302020204030204"},
					Ea:    &XTypeface{Typeface: ""},
					Cs:    &XTypeface{Typeface: ""},
					Font: []*XThemeFont{
						{Script: "Jpan", Typeface: "Yu Gothic Light"},
						{Script: "Hang", Typeface: "맑은 고딕"},
						{Script: "Hans", Typeface: "等线 Light"},
						{Script: "Hant", Typeface: "新細明體"},
						{Script: "Arab", Typeface: "Times New Roman"},
						{Script: "Hebr", Typeface: "Times New Roman"},
						{Script: "Thai", Typeface: "Tahoma"},
						{Script: "Ethi", Typeface: "Nyala"},
						{Script: "Beng", Typeface: "Vrinda"},
						{Script: "Gujr", Typeface: "Shruti"},
						{Script: "Khmr", Typeface: "MoolBoran"},
						{Script: "Knda", Typeface: "Tunga"},
						{Script: "Guru", Typeface: "Raavi"},
						{Script: "Cans", Typeface: "Euphemia"},
						{Script: "Cher", Typeface: "Plantagenet Cherokee"},
						{Script: "Yiii", Typeface: "Microsoft Yi Baiti"},
						{Script: "Tibt", Typeface: "Microsoft Himalaya"},
						{Script: "Thaa", Typeface: "MV Boli"},
						{Script: "Deva", Typeface: "Mangal"},
						{Script: "Telu", Typeface: "Gautami"},
						{Script: "Taml", Typeface: "Latha"},
						{Script: "Syrc", Typeface: "Estrangelo Edessa"},
						{Script: "Orya", Typeface: "Kalinga"},
						{Script: "Mlym", Typeface: "Kartika"},
						{Script: "Laoo", Typeface: "DokChampa"},
						{Script: "Sinh", Typeface: "Iskoola Pota"},
						{Script: "Mong", Typeface: "Mongolian Baiti"},
						{Script: "Viet", Typeface: "Times New Roman"},
						{Script: "Uigh", Typeface: "Microsoft Uighur"},
						{Script: "Geor", Typeface: "Sylfaen"},
					},
				},
				MinorFont: &XFontElement{
					Latin: &XLatin{Typeface: "Calibri", Panose: "020F0502020204030204"},
					Ea:    &XTypeface{Typeface: ""},
					Cs:    &XTypeface{Typeface: ""},
					Font: []*XThemeFont{
						{Script: "Jpan", Typeface: "Yu Gothic"},
						{Script: "Hang", Typeface: "맑은 고딕"},
						{Script: "Hans", Typeface: "等线"},
						{Script: "Hant", Typeface: "新細明體"},
						{Script: "Arab", Typeface: "Arial"},
						{Script: "Hebr", Typeface: "Arial"},
						{Script: "Thai", Typeface: "Tahoma"},
						{Script: "Ethi", Typeface: "Nyala"},
						{Script: "Beng", Typeface: "Vrinda"},
						{Script: "Gujr", Typeface: "Shruti"},
						{Script: "Khmr", Typeface: "DaunPenh"},
						{Script: "Knda", Typeface: "Tunga"},
						{Script: "Guru", Typeface: "Raavi"},
						{Script: "Cans", Typeface: "Euphemia"},
						{Script: "Cher", Typeface: "Plantagenet Cherokee"},
						{Script: "Yiii", Typeface: "Microsoft Yi Baiti"},
						{Script: "Tibt", Typeface: "Microsoft Himalaya"},
						{Script: "Thaa", Typeface: "MV Boli"},
						{Script: "Deva", Typeface: "Mangal"},
						{Script: "Telu", Typeface: "Gautami"},
						{Script: "Taml", Typeface: "Latha"},
						{Script: "Syrc", Typeface: "Estrangelo Edessa"},
						{Script: "Orya", Typeface: "Kalinga"},
						{Script: "Mlym", Typeface: "Kartika"},
						{Script: "Laoo", Typeface: "DokChampa"},
						{Script: "Sinh", Typeface: "Iskoola Pota"},
						{Script: "Mong", Typeface: "Mongolian Baiti"},
						{Script: "Viet", Typeface: "Arial"},
						{Script: "Uigh", Typeface: "Microsoft Uighur"},
						{Script: "Geor", Typeface: "Sylfaen"},
					},
				},
			},
			FmtScheme: &XFmtScheme{
				Name: "Office",
				FillStyleLst: &XFillStyleLst{
					SolidFill: &XSolidFill{
						SchemeClr: &XSchemeClr{Val: "phClr"},
					},
					GradFill: []*XGradFill{
						{
							RotWithShape: "1",
							GsLst: &XGsLst{
								Gs: []*XGs{
									{
										Pos: "0",
										SchemeClr: &XSchemeClr{
											Val:    "phClr",
											LumMod: &XValAttrElement{Val: "110000"},
											SatMod: &XValAttrElement{Val: "105000"},
											Tint:   &XValAttrElement{Val: "67000"},
										},
									}, {
										Pos: "50000",
										SchemeClr: &XSchemeClr{
											Val:    "phClr",
											LumMod: &XValAttrElement{Val: "105000"},
											SatMod: &XValAttrElement{Val: "103000"},
											Tint:   &XValAttrElement{Val: "73000"},
										},
									}, {
										Pos: "100000",
										SchemeClr: &XSchemeClr{
											Val:    "phClr",
											LumMod: &XValAttrElement{Val: "105000"},
											SatMod: &XValAttrElement{Val: "109000"},
											Tint:   &XValAttrElement{Val: "81000"},
										},
									},
								},
							},
							Lin: &XLin{Ang: "5400000", Scaled: "0"},
						}, {
							RotWithShape: "1",
							GsLst: &XGsLst{
								Gs: []*XGs{
									{
										Pos: "0",
										SchemeClr: &XSchemeClr{
											Val:    "phClr",
											SatMod: &XValAttrElement{Val: "103000"},
											LumMod: &XValAttrElement{Val: "102000"},
											Tint:   &XValAttrElement{Val: "94000"},
										},
									}, {
										Pos: "50000",
										SchemeClr: &XSchemeClr{
											Val:    "phClr",
											SatMod: &XValAttrElement{Val: "110000"},
											LumMod: &XValAttrElement{Val: "100000"},
											Shade:  &XValAttrElement{Val: "100000"},
										},
									}, {
										Pos: "100000",
										SchemeClr: &XSchemeClr{
											Val:    "phClr",
											LumMod: &XValAttrElement{Val: "99000"},
											SatMod: &XValAttrElement{Val: "120000"},
											Shade:  &XValAttrElement{Val: "78000"},
										},
									},
								},
							},
							Lin: &XLin{Ang: "5400000", Scaled: "0"},
						},
					},
				},
				LnStyleLst: &XLnStyleLst{
					Ln: []*XLn{
						{
							W:    "6350",
							Cap:  "flat",
							Cmpd: "sng",
							Algn: "ctr",
							SolidFill: &XSolidFill{
								SchemeClr: &XSchemeClr{
									Val: "phClr",
								},
							},
							PrstDash: &XValAttrElement{Val: "solid"},
							Miter:    &XMiter{Lim: "800000"},
						}, {
							W:    "12700",
							Cap:  "flat",
							Cmpd: "sng",
							Algn: "ctr",
							SolidFill: &XSolidFill{
								SchemeClr: &XSchemeClr{
									Val: "phClr",
								},
							},
							PrstDash: &XValAttrElement{Val: "solid"},
							Miter:    &XMiter{Lim: "800000"},
						}, {
							W:    "19050",
							Cap:  "flat",
							Cmpd: "sng",
							Algn: "ctr",
							SolidFill: &XSolidFill{
								SchemeClr: &XSchemeClr{
									Val: "phClr",
								},
							},
							PrstDash: &XValAttrElement{Val: "solid"},
							Miter:    &XMiter{Lim: "800000"},
						},
					},
				},
				EffectStyleLst: &XEffectStyleLst{
					EffectStyle: []*XEffectStyle{
						{EffectLst: &XEffectLst{}},
						{EffectLst: &XEffectLst{}},
						{EffectLst: &XEffectLst{
							OuterShdw: &XOuterShdw{
								BlurRad:      "57150",
								Dist:         "19050",
								Dir:          "5400000",
								Algn:         "ctr",
								RotWithShape: "0",
								SrgbClr: &XSrgbClr{
									Val:   "000000",
									Alpha: &XValAttrElement{Val: "63000"},
								},
							},
						}},
					},
				},
				BgFillStyleLst: &XBgFillStyleLst{
					SolidFill: []*XSolidFill{
						{
							SchemeClr: &XSchemeClr{
								Val: "phClr",
							},
						},
						{
							SchemeClr: &XSchemeClr{
								Val:    "phClr",
								Tint:   &XValAttrElement{Val: "95000"},
								SatMod: &XValAttrElement{Val: "170000"},
							},
						},
					},
					GradFill: &XGradFill{
						RotWithShape: "1",
						GsLst: &XGsLst{
							Gs: []*XGs{
								{
									Pos: "0",
									SchemeClr: &XSchemeClr{
										Val:    "phClr",
										Tint:   &XValAttrElement{Val: "93000"},
										SatMod: &XValAttrElement{Val: "150000"},
										Shade:  &XValAttrElement{Val: "98000"},
										LumMod: &XValAttrElement{Val: "102000"},
									},
								}, {
									Pos: "50000",
									SchemeClr: &XSchemeClr{
										Val:    "phClr",
										Tint:   &XValAttrElement{Val: "98000"},
										SatMod: &XValAttrElement{Val: "130000"},
										Shade:  &XValAttrElement{Val: "90000"},
										LumMod: &XValAttrElement{Val: "103000"},
									},
								}, {
									Pos: "100000",
									SchemeClr: &XSchemeClr{
										Val:    "phClr",
										Shade:  &XValAttrElement{Val: "63000"},
										SatMod: &XValAttrElement{Val: "120000"},
									},
								},
							},
						},
						Lin: &XLin{Ang: "5400000", Scaled: "0"},
					},
				},
			},
		},
		ObjectDefaults:    &XObjectDefaults{},
		ExtraClrSchemeLst: &XExtraClrSchemeLst{},
		ExtLst: &XExtLst{
			Ext: &XExt{
				URI: "{05A4C25C-085E-4340-85A3-A5531E510DB2}",
				ThemeFamily: &XThemeFamily{
					XmlnsThm15: "http://schemas.microsoft.com/office/thememl/2012/main",
					Name:       "Office Theme",
					ID:         "{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}",
					Vid:        "{4A3C46E8-61CC-4603-A589-7422A47A8E4A}",
				},
			},
		},
	}
}
