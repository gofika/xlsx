package xlsx

import (
	"testing"
	"time"

	"github.com/shopspring/decimal"
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
	assert.Equal(t, sheet.Cell(nameCol, titleRow).GetStringValue(), "Name")
	sheet.SetCellValue(valueCol, titleRow, "Score")
	assert.Equal(t, sheet.Cell(valueCol, titleRow).GetStringValue(), "Score")
	sheet.SetCellValue(nameCol, valueRow, "Jason")
	assert.Equal(t, sheet.Cell(nameCol, valueRow).GetStringValue(), "Jason")
	sheet.SetCellValue(valueCol, valueRow, 100)
	assert.True(t, sheet.Cell(valueCol, valueRow).GetFloatValue().Equal(decimal.NewFromInt(100)))

	timeVal := time.Date(1980, 9, 8, 23, 40, 10, 40, time.UTC)
	// SetAxisCellValue example
	sheet.SetAxisCellValue("C3", 200.50)
	sheet.SetAxisCellValue("D3", timeVal)
	assert.Equal(t, sheet.AxisCell("D3").GetTimeValue(), timeVal)
	sheet.SetAxisCellValue("E4", 10*time.Second)
	assert.Equal(t, sheet.AxisCell("E4").GetDurationValue(), 10*time.Second)

	// cell operator method
	sheet.AxisCell("D4").
		SetTimeValue(timeVal).
		SetNumberFormat("yyyy-mm-dd hh:mm:ss")
	sheet.SetColumnWidth("D:D", 20)
	err := f.SaveFile(docPath)
	assert.Nil(t, err)
}
