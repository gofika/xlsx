package xlsx

import (
	"testing"
	"time"

	"github.com/stretchr/testify/assert"
)

func TestNewSheet(t *testing.T) {
	const docPath = "test_docs/two_sheet.xlsx"

	f := NewFile()
	f.NewSheet("Sheet2")
	err := f.SaveFile(docPath)
	assert.Nil(t, err)

	f, err = OpenFile(docPath)
	assert.Nil(t, err)
	sheet := f.OpenSheet("Sheet2")
	assert.NotNil(t, sheet)
}

func TestSetCellValue(t *testing.T) {
	const docPath = "test_docs/set_cell_values.xlsx"
	f := NewFile()

	sheet := f.OpenSheet("Sheet1")
	sheet.SetCellValue(ColumnNumber("A"), 1, "Name")
	sheet.SetCellValue(ColumnNumber("B"), 1, "Score")
	sheet.SetCellValue(ColumnNumber("A"), 2, "Jason")
	sheet.SetCellValue(ColumnNumber("B"), 2, 100)

	sheet.SetCellValue(ColumnNumber("C"), 3, 200.50)
	sheet.SetCellValue(ColumnNumber("D"), 3, time.Date(1980, 9, 8, 23, 40, 10, 40, time.Local))
	sheet.SetCellValue(ColumnNumber("E"), 4, 10*time.Second)

	sheet.AxisCell("D4").
		SetTimeValue(time.Date(1980, 9, 8, 23, 40, 10, 40, time.Local)).
		SetNumberFormat("yyyy-mm-dd hh:mm:ss")
	err := f.SaveFile(docPath)
	assert.Nil(t, err)
}
