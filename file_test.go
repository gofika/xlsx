package xlsx

import (
	"os"
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestNewFile(t *testing.T) {
	f := NewFile()
	err := os.MkdirAll("test_docs", os.ModePerm)
	assert.Nil(t, err)
	err = f.SaveFile("test_docs/empty.xlsx")
	assert.Nil(t, err)
}

func TestNewFileSheet(t *testing.T) {
	const customSheetName = "MySheet"
	f := NewFile(WithDefaultSheetName(customSheetName))
	sheet := f.OpenSheet(customSheetName)
	assert.NotNil(t, sheet)
}

func TestOpenFile(t *testing.T) {
	f, err := OpenFile("test_docs/empty.xlsx")
	assert.Nil(t, err)
	assert.Len(t, f.Sheets(), 1)
}
