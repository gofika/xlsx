package packaging

import (
	"encoding/xml"
	"fmt"
)

// Workbook Defines
const (
	WorkbookContentType      = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
	WorkbookRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"

	WorkbookPath     = "xl"
	WorkbookFileName = "workbook.xml"
)

// XWorkbook Workbook XML struct
type XWorkbook struct {
	XMLName     xml.Name `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main workbook"`
	XmlnsR      string   `xml:"xmlns:r,attr"`
	XmlnsMC     string   `xml:"xmlns:mc,attr"`
	McIgnorable string   `xml:"mc:Ignorable,attr"`
	XmlnsX15    string   `xml:"xmlns:x15,attr"`

	FileVersion *XFileVersion `xml:"fileVersion"`
	WorkbookPr  *XWorkbookPr  `xml:"workbookPr"`
	BookViews   *XBookViews   `xml:"bookViews"`
	Sheets      *XSheets      `xml:"sheets"`
	CalcPr      *XCalcPr      `xml:"calcPr"`
}

// XFileVersion FileVersion type
type XFileVersion struct {
	AppName      string `xml:"appName,attr,omitempty"`
	LastEdited   string `xml:"lastEdited,attr,omitempty"`
	LowestEdited string `xml:"lowestEdited,attr,omitempty"`
	RupBuild     string `xml:"rupBuild,attr,omitempty"`
}

// XWorkbookPr WorkbookPr type
type XWorkbookPr struct {
	DefaultThemeVersion string `xml:"defaultThemeVersion,attr,omitempty"`
}

// XBookViews BookViews type
type XBookViews struct {
	WorkBookView []*XWorkBookView `xml:"workbookView"`
}

// XWorkBookView WorkBookView type
type XWorkBookView struct {
	XWindow      string `xml:"xWindow,attr,omitempty"`
	YWindow      string `xml:"yWindow,attr,omitempty"`
	WindowWidth  int    `xml:"windowWidth,attr,omitempty"`
	WindowHeight int    `xml:"windowHeight,attr,omitempty"`
}

// XSheets Sheets type
type XSheets struct {
	Sheet []*XSheet `xml:"sheet"`
}

// XSheet Sheet type
type XSheet struct {
	Name    string `xml:"name,attr,omitempty"`
	SheetID int    `xml:"sheetId,attr,omitempty"`
	Rid     string `xml:"r:id,attr,omitempty"`
}

// XCalcPr CalcPr type
type XCalcPr struct {
	CalcID string `xml:"calcId,attr,omitempty"`
}

// NewXWorkbook create *XWorkbook from WorksheetRelationships
func NewXWorkbook(worksheetRelations *XRelationships) (workbook *XWorkbook) {
	workbook = &XWorkbook{
		XmlnsR:      "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
		XmlnsMC:     "http://schemas.openxmlformats.org/markup-compatibility/2006",
		McIgnorable: "x15",
		XmlnsX15:    "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main",
		FileVersion: &XFileVersion{
			AppName:      "xl",
			LastEdited:   "6",
			LowestEdited: "6",
			RupBuild:     "14420",
		},
		WorkbookPr: &XWorkbookPr{
			DefaultThemeVersion: "164011",
		},
		BookViews: &XBookViews{
			WorkBookView: []*XWorkBookView{
				{
					XWindow:      "0",
					YWindow:      "0",
					WindowWidth:  22260,
					WindowHeight: 12645,
				},
			},
		},
		Sheets: &XSheets{Sheet: []*XSheet{}},
		CalcPr: &XCalcPr{
			CalcID: "162913",
		},
	}
	for _, relationship := range worksheetRelations.Relationships {
		if relationship.Type != WorksheetRelationshipType {
			continue
		}
		workbook.Sheets.Sheet = append(workbook.Sheets.Sheet, &XSheet{
			Name:    fmt.Sprintf("Sheet%d", relationship.Index),
			SheetID: relationship.Index,
			Rid:     relationship.ID,
		})
	}

	return
}
