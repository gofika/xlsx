package packaging

import (
	"path"
	"testing"

	"github.com/gofika/util/fileutil"
	"github.com/stretchr/testify/assert"
)

const defaultSharedStringsTemplate = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"></sst>`

func testNewXSharedStrings(t *testing.T, file *XFile) {
	var result string
	var err error
	result, err = XMLMarshalAppendHeadIndent(file.SharedStrings)
	assert.Nil(t, err)
	if needWriteTestFile {
		err = fileutil.WriteFile(path.Join(templatePath, SharedStringsPath, SharedStringsFileName), []byte(result))
		assert.Nil(t, err)
	}
	assert.Equal(t, result, defaultSharedStringsTemplate)
}
