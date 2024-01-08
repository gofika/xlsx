package packaging

import (
	"path"
	"testing"

	"github.com/gofika/util/fileutil"
	"github.com/stretchr/testify/assert"
)

const defaultStyleSheetTemplate = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main">
    <fonts count="1" x14ac:knownFonts="1">
        <font>
            <sz val="11"></sz>
            <color theme="1"></color>
            <name val="Microsoft YaHei"></name>
            <family val="0"></family>
            <charset val="0"></charset>
        </font>
    </fonts>
    <fills count="2">
        <fill>
            <patternFill patternType="none"></patternFill>
        </fill>
        <fill>
            <patternFill patternType="gray125"></patternFill>
        </fill>
    </fills>
    <borders count="1">
        <border>
            <left></left>
            <right></right>
            <top></top>
            <bottom></bottom>
            <diagonal></diagonal>
        </border>
    </borders>
    <cellStyleXfs count="1">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"></xf>
    </cellStyleXfs>
    <cellXfs count="1">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"></xf>
    </cellXfs>
    <cellStyles count="1">
        <cellStyle name="Normal" xfId="0" builtinId="0"></cellStyle>
    </cellStyles>
    <dxfs count="0"></dxfs>
    <tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"></tableStyles>
    <extLst>
        <ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
            <x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"></x14:slicerStyles>
        </ext>
        <ext uri="{9260A510-F301-46a8-8635-F512D64BE5F5}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
            <x15:timelineStyles defaultTimelineStyle="TimeSlicerStyleLight1"></x15:timelineStyles>
        </ext>
    </extLst>
</styleSheet>`

func testNewDefaultXStyleSheet(t *testing.T, file *XFile) {
	var result string
	var err error
	result, err = XMLMarshalAppendHeadIndent(file.StyleSheet)
	assert.Nil(t, err)
	if needWriteTestFile {
		err = fileutil.WriteFile(path.Join(templatePath, StyleSheetPath, StyleSheetFileName), []byte(result))
		assert.Nil(t, err)
	}
	assert.Equal(t, result, defaultStyleSheetTemplate)
}
