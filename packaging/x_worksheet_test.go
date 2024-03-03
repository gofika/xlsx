package packaging

import (
	"fmt"
	"path"
	"testing"

	"github.com/gofika/util/fileutil"
	"github.com/stretchr/testify/assert"
)

const defaultWorksheetTemplate = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
    <dimension ref="A1"></dimension>
    <sheetViews>
        <sheetView tabSelected="1" workbookViewId="0"></sheetView>
    </sheetViews>
    <sheetFormatPr defaultRowHeight="15"></sheetFormatPr>
    <sheetData></sheetData>
    <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"></pageMargins>
</worksheet>`

func testNewDefaultXWorksheet(t *testing.T, file *XFile) {
	var result string
	var err error
	result, err = XMLMarshalAppendHeadIndent(file.Worksheets[0])
	assert.Nil(t, err)
	if needWriteTestFile {
		err = fileutil.WriteFile(path.Join(templatePath, WorksheetPath, fmt.Sprintf(WorksheetFileName, 1)), []byte(result))
		assert.Nil(t, err)
	}
	assert.Equal(t, result, defaultWorksheetTemplate)
}
