package packaging

import (
	"testing"
)

func TestFile(t *testing.T) {
	file := NewDefaultFile("Microsoft YaHei", 11)
	testTheme(t, file)
	testNewXContentTypes(t, file)
	testNewDefaultXCoreProperties(t, file)
	testNewXExtendedProperties(t, file)
	testNewWorkbookXRelationships(t, file)
	testNewDefaultRootXRelationships(t, file)
	testNewXWorkbook(t, file)
	testNewDefaultXWorksheet(t, file)
	testNewDefaultXStyleSheet(t, file)
	testNewXSharedStrings(t, file)
}
