package packaging

import (
	"encoding/xml"
	"fmt"
	"path"
)

// Relationships Defines
const (
	WorkbookRelationshipsPath     = "xl/_rels"
	WorkbookRelationshipsFileName = "workbook.xml.rels"

	RootRelationshipsPath     = "_rels"
	RootRelationshipsFileName = ".rels"
)

// XRelationships .rels XMLDocument type
type XRelationships struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/package/2006/relationships Relationships"`

	Relationships []*XRelationship `xml:"Relationship"`
}

// XRelationship Relationship type
type XRelationship struct {
	ID     string `xml:"Id,attr"`
	Type   string `xml:"Type,attr"`
	Target string `xml:"Target,attr"`
	Index  int    `xml:"-"`
}

// NewWorkbookXRelationships create *XRelationships from Worksheets and Themes
func NewWorkbookXRelationships(file *XFile) (workbookRelations *XRelationships) {
	workbookRelations = &XRelationships{
		Relationships: []*XRelationship{},
	}
	rID := 0

	// worksheets/sheet?.xml
	for i := range file.Worksheets {
		sheetIndex := i + 1
		rID++
		workbookRelations.Relationships = append(workbookRelations.Relationships, &XRelationship{
			ID:     fmt.Sprintf("rId%d", rID),
			Type:   WorksheetRelationshipType,
			Target: fmt.Sprintf(WorksheetFileName, sheetIndex),
			Index:  rID,
		})
	}

	// theme/theme?.xml
	for i := range file.Themes {
		themeIndex := i + 1
		rID++
		workbookRelations.Relationships = append(workbookRelations.Relationships, &XRelationship{
			ID:     fmt.Sprintf("rId%d", rID),
			Type:   ThemeRelationshipType,
			Target: fmt.Sprintf(ThemeFileName, themeIndex),
			Index:  rID,
		})
	}

	// styles.xml
	rID++
	workbookRelations.Relationships = append(workbookRelations.Relationships, &XRelationship{
		ID:     fmt.Sprintf("rId%d", rID),
		Type:   StyleSheetRelationshipType,
		Target: StyleSheetFileName,
		Index:  rID,
	})

	// sharedStrings.xml
	if file.SharedStrings.Count > 0 {
		rID++
		workbookRelations.Relationships = append(workbookRelations.Relationships, &XRelationship{
			ID:     fmt.Sprintf("rId%d", rID),
			Type:   SharedStringsRelationshipType,
			Target: SharedStringsFileName,
			Index:  rID,
		})
	}

	return
}

// NewDefaultRootXRelationships create *XRelationships with default template
func NewDefaultRootXRelationships() *XRelationships {
	return &XRelationships{
		Relationships: []*XRelationship{
			{
				ID:     "rId1",
				Type:   WorkbookRelationshipType,
				Target: path.Join(WorkbookPath, WorkbookFileName),
				Index:  1,
			},
			{
				ID:     "rId2",
				Type:   CorePropertiesRelationshipType,
				Target: path.Join(CorePropertiesPath, CorePropertiesFileName),
				Index:  2,
			},
			{
				ID:     "rId3",
				Type:   ExtendedPropertiesRelationshipType,
				Target: path.Join(ExtendedPropertiesPath, ExtendedPropertiesFileName),
				Index:  3,
			},
		},
	}
}
