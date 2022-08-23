package packaging

import (
	"path"
	"testing"

	"github.com/gofika/util/fileutil"
	"github.com/stretchr/testify/assert"
)

const defaultCorePropertiesTemplate = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
    <dc:creator>Microsoft</dc:creator>
    <cp:lastModifiedBy></cp:lastModifiedBy>
    <dcterms:created xsi:type="dcterms:W3CDTF">2015-06-05T18:19:34Z</dcterms:created>
    <dcterms:modified xsi:type="dcterms:W3CDTF">2015-06-05T18:19:39Z</dcterms:modified>
</cp:coreProperties>`

func testNewDefaultXCoreProperties(t *testing.T, file *XFile) {
	var result string
	var err error
	result, err = XMLMarshalAppendHeadIndent(file.CoreProperties)
	assert.Nil(t, err)
	if needWriteTestFile {
		err = fileutil.WriteFile(path.Join(templatePath, CorePropertiesPath, CorePropertiesFileName), []byte(result))
		assert.Nil(t, err)
	}
	assert.Equal(t, result, defaultCorePropertiesTemplate)
}
