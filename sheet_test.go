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

	// table example
	const titleRow = 1
	const valueRow = 2
	nameCol := ColumnNumber("A")
	valueCol := ColumnNumber("B")
	sheet.SetCellValue(nameCol, titleRow, "Name")
	sheet.SetCellValue(valueCol, titleRow, "Score")
	sheet.SetCellValue(nameCol, valueRow, "Jason")
	sheet.SetCellValue(valueCol, valueRow, 100)

	// SetAxisCellValue example
	sheet.SetAxisCellValue("C3", 200.50)
	sheet.SetAxisCellValue("D3", time.Date(1980, 9, 8, 23, 40, 10, 40, time.UTC))
	sheet.SetAxisCellValue("E4", 10*time.Second)

	// cell operator method
	sheet.AxisCell("D4").
		SetTimeValue(time.Date(1980, 9, 8, 23, 40, 10, 40, time.UTC)).
		SetNumberFormat("yyyy-mm-dd hh:mm:ss")
	err := f.SaveFile(docPath)
	assert.Nil(t, err)
}
