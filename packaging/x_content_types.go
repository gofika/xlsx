package packaging

import (
	"encoding/xml"
	"fmt"
	"path"
)

// ContentTypes Defines
const (
	ContentTypesPath     = ""
	ContentTypesFileName = "[Content_Types].xml"
)

// XContentTypes directly maps the types element of content types for relationships
type XContentTypes struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/package/2006/content-types Types"`

	Defaults  []*XDefault  `xml:"Default"`
	Overrides []*XOverride `xml:"Override"`
}

// XDefault directly maps the override element in the namespace
type XDefault struct {
	Extension   string `xml:",attr"`
	ContentType string `xml:",attr"`
}

// XOverride directly maps the default element in the namespace
type XOverride struct {
	PartName    string `xml:",attr"`
	ContentType string `xml:",attr"`
}

// NewXContentTypes create XContentTypes from WorksheetRelations
func NewXContentTypes(worksheetRelations *XRelationships) (contentTypes *XContentTypes) {
	contentTypes = &XContentTypes{
		Defaults: []*XDefault{
			//{
			//	Extension:   "bin",
			//	ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings",
			//},
			{
				Extension:   "rels",
				ContentType: RelationshipContentType,
			},
			{
				Extension:   "xml",
				ContentType: "application/xml",
			},
		},
		Overrides: []*XOverride{
			{
				PartName:    path.Join("/", WorkbookPath, WorkbookFileName),
				ContentType: WorkbookContentType,
			},
			{
				PartName:    path.Join("/", CorePropertiesPath, CorePropertiesFileName),
				ContentType: CorePropertiesContentType,
			},
			{
				PartName:    path.Join("/", ExtendedPropertiesPath, ExtendedPropertiesFileName),
				ContentType: ExtendedPropertiesContentType,
			},
		},
	}

	for _, relationship := range worksheetRelations.Relationships {
		contentType := ""
		switch relationship.Type {
		case StyleSheetRelationshipType:
			contentType = StyleSheetContentType
		case ThemeRelationshipType:
			contentType = ThemeContentType
		case WorksheetRelationshipType:
			contentType = WorksheetContentType
		case SharedStringsRelationshipType:
			contentType = SharedStringsContentType
		}
		contentTypes.Overrides = append(contentTypes.Overrides, &XOverride{
			PartName:    fmt.Sprintf("/xl/%s", relationship.Target),
			ContentType: contentType,
		})
	}
	return
}
