package xlsx

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestColumnTitle(t *testing.T) {
	assert.Equal(t, ColumnName(0), "")
	assert.Equal(t, ColumnName(1), "A")
	assert.Equal(t, ColumnName(26), "Z")
	assert.Equal(t, ColumnName(26*2+1), "BA")
	assert.Equal(t, ColumnName(26*3+1), "CA")
	assert.Equal(t, ColumnName(26*26+1), "ZA")
}

func TestColumnNumber(t *testing.T) {
	assert.Equal(t, ColumnNumber("WrongNumber"), 0)
	assert.Equal(t, ColumnNumber("A"), 1)
	assert.Equal(t, ColumnNumber("Z"), 26)
	assert.Equal(t, ColumnNumber("AA"), 26+1)
	assert.Equal(t, ColumnNumber("BA"), 26*2+1)
	assert.Equal(t, ColumnNumber("CA"), 26*3+1)
	assert.Equal(t, ColumnNumber("ZA"), 26*26+1)
}

func TestCellNameToCoordinates(t *testing.T) {
	var col, row int
	col, row = CellNameToCoordinates("A5")
	assert.Equal(t, col, 1)
	assert.Equal(t, row, 5)

	col, row = CellNameToCoordinates("Z9")
	assert.Equal(t, col, 26)
	assert.Equal(t, row, 9)
}
