package xlsx

import (
	"archive/zip"
	"compress/flate"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"path"
	"regexp"
	"strconv"

	"github.com/gofika/util/gobutil"
	"github.com/gofika/xlsx/packaging"
)

// fileImpl define for operation xlsx file
type fileImpl struct {
	xFile *packaging.XFile
}

func newFile() *fileImpl {
	return &fileImpl{
		xFile: packaging.NewDefaultFile("Microsoft YaHei", 11),
	}
}

func newFileWithFont(defaultFontName string, defaultFontSize int) *fileImpl {
	return &fileImpl{
		xFile: packaging.NewDefaultFile(defaultFontName, defaultFontSize),
	}
}

// openFile open a xlsx *fileImpl
func openFile(name string) (*fileImpl, error) {
	file := &fileImpl{
		xFile: &packaging.XFile{
			ContentTypes:          &packaging.XContentTypes{},
			Worksheets:            []*packaging.XWorksheet{},
			Workbook:              &packaging.XWorkbook{},
			WorkbookRelationships: &packaging.XRelationships{},
			RootRelationships:     &packaging.XRelationships{},
			ExtendedProperties:    &packaging.XExtendedProperties{},
			CoreProperties:        &packaging.XCoreProperties{},
			Themes:                []*packaging.XTheme{},
			StyleSheet:            &packaging.XStyleSheet{},
			SharedStrings:         &packaging.XSharedStrings{},
		},
	}
	r, err := zip.OpenReader(name)
	if err != nil {
		return nil, err
	}
	err = file.readParts(&r.Reader)
	_ = r.Close()
	if err != nil {
		return nil, err
	}
	return file, nil
}

// SaveFile save xlsx file
func (f *fileImpl) SaveFile(name string) error {
	file, err := os.Create(name)
	if err != nil {
		return nil
	}
	err = f.Write(file)
	errClose := file.Close()
	if err != nil {
		return err
	}
	return errClose
}

// Write save to steam
func (f *fileImpl) Write(w io.Writer) error {
	zipWriter := zip.NewWriter(w)
	zipWriter.RegisterCompressor(zip.Deflate, func(out io.Writer) (io.WriteCloser, error) {
		return flate.NewWriter(out, flate.BestCompression)
	})
	err := f.writeParts(zipWriter)
	errClose := zipWriter.Close()
	if err != nil {
		return err
	}
	return errClose
}

func writePart(zipWriter *zip.Writer, name string, part any) error {
	w, err := zipWriter.Create(name)
	if err != nil {
		return err
	}
	_, err = w.Write([]byte(packaging.XMLProlog))
	if err != nil {
		return err
	}
	xmlWriter := xml.NewEncoder(w)
	// for easy look
	xmlWriter.Indent("", "    ")

	err = xmlWriter.Encode(part)
	if err != nil {
		return err
	}
	return nil
}

func (f *fileImpl) writeParts(zipWriter *zip.Writer) error {
	// RootRelationships _rels/.rels
	err := writePart(zipWriter, path.Join(packaging.RootRelationshipsPath, packaging.RootRelationshipsFileName), f.xFile.RootRelationships)
	if err != nil {
		return err
	}
	// Workbook xl/workbook.xml
	err = writePart(zipWriter, path.Join(packaging.WorkbookPath, packaging.WorkbookFileName), f.xFile.Workbook)
	if err != nil {
		return err
	}
	// WorkbookRelationships xl/_rels/workbook.xml.rels
	err = writePart(zipWriter, path.Join(packaging.WorkbookRelationshipsPath, packaging.WorkbookRelationshipsFileName), f.xFile.WorkbookRelationships)
	if err != nil {
		return err
	}
	// CoreProperties docProps/core.xml
	err = writePart(zipWriter, path.Join(packaging.CorePropertiesPath, packaging.CorePropertiesFileName), f.xFile.CoreProperties)
	if err != nil {
		return err
	}
	// ExtendedProperties docProps/app.xml
	err = writePart(zipWriter, path.Join(packaging.ExtendedPropertiesPath, packaging.ExtendedPropertiesFileName), f.xFile.ExtendedProperties)
	if err != nil {
		return err
	}
	// Worksheets xl/worksheets/sheet?.xml
	for i, worksheet := range f.xFile.Worksheets {
		err = writePart(zipWriter, path.Join(packaging.WorksheetPath, fmt.Sprintf(packaging.WorksheetFileName, i+1)), worksheet)
		if err != nil {
			return err
		}
	}
	// Themes xl/theme/theme?.xml
	for i, theme := range f.xFile.Themes {
		err = writePart(zipWriter, path.Join(packaging.ThemePath, fmt.Sprintf(packaging.ThemeFileName, i+1)), theme)
		if err != nil {
			return err
		}
	}
	// StyleSheet xl/styles.xml
	err = writePart(zipWriter, path.Join(packaging.StyleSheetPath, packaging.StyleSheetFileName), f.xFile.StyleSheet)
	if err != nil {
		return err
	}
	// SharedStrings xl/sharedStrings.xml
	if f.xFile.SharedStrings != nil && f.xFile.SharedStrings.Count > 0 {
		err = writePart(zipWriter, path.Join(packaging.SharedStringsPath, packaging.SharedStringsFileName), f.xFile.SharedStrings)
		if err != nil {
			return err
		}
	}
	// ContentTypes [Content_Types].xml
	err = writePart(zipWriter, path.Join(packaging.ContentTypesPath, packaging.ContentTypesFileName), f.xFile.ContentTypes)
	if err != nil {
		return err
	}
	return nil
}

func readPart(zipReader *zip.Reader, name string, part any) error {
	var zFile *zip.File
	for _, z := range zipReader.File {
		if z.Name == name {
			zFile = z
			break
		}
	}
	if zFile == nil {
		return fmt.Errorf("readPart file %s not exist", name)
	}
	r, err := zFile.Open()
	if err != nil {
		return err
	}
	xmlReader := xml.NewDecoder(r)
	return xmlReader.Decode(part)
}

// readRootRelationships load from _rels/.rels
func readRootRelationships(zipReader *zip.Reader, xFile *packaging.XFile) error {
	err := readPart(zipReader, path.Join(packaging.RootRelationshipsPath, packaging.RootRelationshipsFileName), xFile.RootRelationships)
	if err != nil {
		return err
	}
	for _, relationship := range xFile.RootRelationships.Relationships {
		switch relationship.Type {
		case packaging.WorkbookRelationshipType: // xl/workbook.xml
			var workbookU packaging.XWorkbookU // fix xml ns bug
			err = readPart(zipReader, relationship.Target, &workbookU)
			if err != nil {
				return err
			}
			err = gobutil.DeepCopy(xFile.Workbook, &workbookU)
			if err != nil {
				return err
			}
			// xl/_rels/workbook.xml.rels
			err = readPart(zipReader, path.Join(packaging.WorkbookRelationshipsPath, packaging.WorkbookRelationshipsFileName), xFile.WorkbookRelationships)
			if err != nil {
				return err
			}
		case packaging.CorePropertiesRelationshipType: // docProps/core.xml
			var corePropertiesU packaging.XCorePropertiesU // fix xml ns bug
			err := readPart(zipReader, relationship.Target, &corePropertiesU)
			if err != nil {
				return err
			}
			err = gobutil.DeepCopy(xFile.CoreProperties, &corePropertiesU)
			if err != nil {
				return err
			}
		case packaging.ExtendedPropertiesRelationshipType: // docProps/app.xml
			var extendedPropertiesU packaging.XExtendedPropertiesU // fix xml ns bug
			err := readPart(zipReader, relationship.Target, &extendedPropertiesU)
			if err != nil {
				return err
			}
			err = gobutil.DeepCopy(xFile.ExtendedProperties, &extendedPropertiesU)
			if err != nil {
				return err
			}
		default:
			return fmt.Errorf("readParts: unknown relationship type %s", relationship.Type)
		}
	}
	return nil
}

// readWorkbookRelationships load from xl/_rels/workbook.xml.rels
func readWorkbookRelationships(zipReader *zip.Reader, xFile *packaging.XFile) (err error) {
	for _, relationship := range xFile.WorkbookRelationships.Relationships {
		switch relationship.Type {
		case packaging.WorksheetRelationshipType: // worksheets/sheet?.xml
			re, _ := regexp.Compile(`worksheets/sheet([1-9][0-9]*).xml`)
			mcs := re.FindStringSubmatch(relationship.Target)
			if len(mcs) < 2 {
				return fmt.Errorf("readParts: unknown sheet target exp %s", relationship.Target)
			}
			sheetIndex, _ := strconv.Atoi(mcs[1])
			relationship.Index = sheetIndex
			for len(xFile.Worksheets) < sheetIndex {
				xFile.Worksheets = append(xFile.Worksheets, &packaging.XWorksheet{})
			}
			err = readPart(zipReader, path.Join("xl", relationship.Target), xFile.Worksheets[sheetIndex-1])
			if err != nil {
				return
			}
		case packaging.ThemeRelationshipType: // theme/theme?.xml
			re, _ := regexp.Compile(`theme/theme([1-9][0-9]*).xml`)
			mcs := re.FindStringSubmatch(relationship.Target)
			if len(mcs) < 2 {
				return fmt.Errorf("readParts: unknown theme target exp %s", relationship.Target)
			}
			themeIndex, _ := strconv.Atoi(mcs[1])
			relationship.Index = themeIndex
			for len(xFile.Themes) < themeIndex {
				xFile.Themes = append(xFile.Themes, &packaging.XTheme{})
			}
			var themeU packaging.XThemeU // fix xml ns bug
			err = readPart(zipReader, path.Join("xl", relationship.Target), &themeU)
			if err != nil {
				return
			}
			err = gobutil.DeepCopy(xFile.Themes[themeIndex-1], &themeU)
			if err != nil {
				return
			}
		case packaging.StyleSheetRelationshipType: // styles.xml
			var styleSheetU packaging.XStyleSheetU
			err = readPart(zipReader, path.Join("xl", relationship.Target), &styleSheetU)
			if err != nil {
				return
			}
			err = gobutil.DeepCopy(xFile.StyleSheet, &styleSheetU)
			if err != nil {
				return
			}
		case packaging.SharedStringsRelationshipType: // sharedStrings.xml
			err = readPart(zipReader, path.Join("xl", relationship.Target), &xFile.SharedStrings)
			if err != nil {
				return
			}
		default:
			return fmt.Errorf("[xl/_rels/workbook.xml.rels] Unknown Relationship Type: %s", relationship.Type)
		}
	}
	return nil
}

// readParts load from *zip.Reader
func (f *fileImpl) readParts(zipReader *zip.Reader) (err error) {
	// _rels/.rels
	err = readRootRelationships(zipReader, f.xFile)
	if err != nil {
		return
	}
	// xl/_rels/workbook.xml.rels
	err = readWorkbookRelationships(zipReader, f.xFile)
	if err != nil {
		return
	}
	// [Content_Types].xml
	err = readPart(zipReader, path.Join(packaging.ContentTypesPath, packaging.ContentTypesFileName), f.xFile.ContentTypes)
	if err != nil {
		return
	}
	return
}

// OpenSheet open a exist *sheetImpl by name
//
// Example:
//
//	sheet := file.OpenSheet("Sheet1")
//
// return nil if sheet not exist
func (f *fileImpl) OpenSheet(name string) Sheet {
	name = trimSheetName(name)
	for i, sheet := range f.xFile.Workbook.Sheets.Sheet {
		if sheet.Name == name {
			return newSheet(f, sheet, i)
		}
	}
	return nil
}

func (f *fileImpl) getWorkbook() *packaging.XWorkbook {
	return f.xFile.Workbook
}

func (f *fileImpl) getExtendedProperties() *packaging.XExtendedProperties {
	return f.xFile.ExtendedProperties
}

// NewSheet create a new *sheetImpl with sheet name
// Example:
//
//	sheet := file.NewSheet("Sheet2")
func (f *fileImpl) NewSheet(name string) Sheet {
	name = trimSheetName(name)
	if f.OpenSheet(name) != nil {
		return nil
	}
	// append worksheet
	sheetID := 0
	workbook := f.getWorkbook()
	sheetIndex := len(workbook.Sheets.Sheet)
	for _, sheet := range workbook.Sheets.Sheet {
		if sheet.SheetID > sheetID {
			sheetID = sheet.SheetID
		}
	}
	sheetID++ // new sheetId
	worksheet := packaging.NewDefaultXWorksheet()
	f.xFile.Worksheets = append(f.xFile.Worksheets, worksheet)

	// update workbook.xml.rels
	f.xFile.WorkbookRelationships = packaging.NewWorkbookXRelationships(f.xFile)
	rID := 0
	for _, i := range f.xFile.WorkbookRelationships.Relationships {
		if i.Type != packaging.WorksheetRelationshipType {
			continue
		}
		id, _ := strconv.Atoi(i.ID[3:])
		if id > rID {
			rID = id
		}
	}

	// update workbook.xml
	sheet := &packaging.XSheet{
		Name:    name,
		SheetID: sheetID,
		Rid:     fmt.Sprintf("rId%d", rID),
	}
	workbook.Sheets.Sheet = append(workbook.Sheets.Sheet, sheet)

	// update docProps/app.xml
	extendedProperties := f.getExtendedProperties()
	headingPairsVector := newVector(extendedProperties.HeadingPairs.Vector)
	headingPairs := headingPairsVector.GetIntPairs()
	headingPairs["Worksheets"]++
	headingPairsVector.SetIntPairs(headingPairs)

	titlesOfPartsVector := newVector(extendedProperties.TitlesOfParts.Vector)
	titlesOfPartsVector.AppendString(name)

	// update [Content_Types].xml
	f.xFile.ContentTypes = packaging.NewXContentTypes(f.xFile.WorkbookRelationships)

	return newSheet(f, sheet, sheetIndex)
}

func (f *fileImpl) Sheets() (sheets []Sheet) {
	sheets = []Sheet{}
	for i, sheet := range f.xFile.Workbook.Sheets.Sheet {
		sheets = append(sheets, newSheet(f, sheet, i))
	}
	return
}
