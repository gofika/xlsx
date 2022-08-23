package xlsx

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestNewFile(t *testing.T) {
	f := NewFile()
	err := f.SaveFile("test_docs/empty.xlsx")
	assert.Nil(t, err)
}

func TestOpenFile(t *testing.T) {
	f, err := OpenFile("test_docs/two_sheet.xlsx")
	assert.Nil(t, err)
	assert.Len(t, f.Sheets(), 2)
}
