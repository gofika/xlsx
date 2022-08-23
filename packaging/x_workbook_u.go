package packaging

import (
	"encoding/xml"
)

// XWorkbookU fix XML ns
type XWorkbookU struct {
	XMLName     xml.Name `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main workbook"`
	XmlnsR      string   `xml:"r,attr"`
	XmlnsMC     string   `xml:"mc,attr"`
	McIgnorable string   `xml:"Ignorable,attr"`
	XmlnsX15    string   `xml:"x15,attr"`

	FileVersion *XFileVersionU `xml:"fileVersion"`
	WorkbookPr  *XWorkbookPrU  `xml:"workbookPr"`
	BookViews   *XBookViewsU   `xml:"bookViews"`
	Sheets      *XSheetsU      `xml:"sheets"`
	CalcPr      *XCalcPrU      `xml:"calcPr"`
}

// XFileVersionU fix XML ns
type XFileVersionU struct {
	AppName      string `xml:"appName,attr,omitempty"`
	LastEdited   string `xml:"lastEdited,attr,omitempty"`
	LowestEdited string `xml:"lowestEdited,attr,omitempty"`
	RupBuild     string `xml:"rupBuild,attr,omitempty"`
}

// XWorkbookPrU fix XML ns
type XWorkbookPrU struct {
	DefaultThemeVersion string `xml:"defaultThemeVersion,attr,omitempty"`
}

// XBookViewsU fix XML ns
type XBookViewsU struct {
	WorkBookView []*XWorkBookViewU `xml:"workbookView"`
}

// XWorkBookViewU fix XML ns
type XWorkBookViewU struct {
	XWindow      string `xml:"xWindow,attr,omitempty"`
	YWindow      string `xml:"yWindow,attr,omitempty"`
	WindowWidth  int    `xml:"windowWidth,attr,omitempty"`
	WindowHeight int    `xml:"windowHeight,attr,omitempty"`
}

// XSheetsU fix XML ns
type XSheetsU struct {
	Sheet []*XSheetU `xml:"sheet"`
}

// XSheetU fix XML ns
type XSheetU struct {
	Name    string `xml:"name,attr,omitempty"`
	SheetID int    `xml:"sheetId,attr,omitempty"`
	Rid     string `xml:"id,attr,omitempty"`
}

// XCalcPrU fix XML ns
type XCalcPrU struct {
	CalcID string `xml:"calcId,attr,omitempty"`
}
