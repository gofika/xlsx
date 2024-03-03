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
	f := NewFile(WithDefaultFontSize(9))

	sheet := f.OpenSheet("Sheet1")

	// table example
	const titleRow = 1
	const valueRow = 2
	nameCol := ColumnNumber("A")
	valueCol := ColumnNumber("B")
	style := sheet.SetCellValue(nameCol, titleRow, "Name").GetStyle()
	assert.Equal(t, sheet.Cell(nameCol, titleRow).GetStringValue(), "Name")
	// set border style
	style.Border.BottomBorder = BorderStyleThin
	style.Border.BottomBorderColor = Color{
		Color: "FF0000",
	}
	// set cell alignment
	style.Alignment.Horizontal = HorizontalAlignmentCenter
	style.Alignment.Vertical = VerticalAlignmentCenter
	// set font style
	style.Font.Bold = true
	sheet.SetCellStyle(nameCol, titleRow, style)
	sheet.SetCellValue(valueCol, titleRow, "Score")
	assert.Equal(t, sheet.Cell(valueCol, titleRow).GetStringValue(), "Score")
	sheet.SetCellStyle(valueCol, titleRow, style)
	sheet.SetCellValue(nameCol, valueRow, "Jason")
	assert.Equal(t, sheet.Cell(nameCol, valueRow).GetStringValue(), "Jason")
	sheet.SetCellValue(valueCol, valueRow, 100)
	assert.True(t, sheet.Cell(valueCol, valueRow).GetFloatValue().Equal(decimal.NewFromInt(100)))

	timeVal := time.Date(1980, 9, 8, 23, 40, 10, 40, time.UTC)
	// merge cell
	sheet.MergeCell("D3", "E3")
	// SetAxisCellValue example
	sheet.SetAxisCellValue("C3", 200.50)
	sheet.SetAxisCellValue("D3", timeVal)
	assert.Equal(t, sheet.AxisCell("D3").GetTimeValue(), timeVal)
	sheet.SetAxisCellValue("E4", 10*time.Second)
	assert.Equal(t, sheet.AxisCell("E4").GetDurationValue(), 10*time.Second)
	// SetCellBorder example
	sheet.SetCellBorder(3, 3, BorderStyleThin, Color{Color: "00FF00"}, false, true, false, true)
	sheet.SetAxisCellBorder("E4", BorderStyleThin, Color{Color: "00FF00"}, false, true, false, true)

	// cell operator method
	sheet.AxisCell("D4").
		SetTimeValue(timeVal).
		SetNumberFormat("yyyy-mm-dd hh:mm:ss")
	sheet.SetColumnWidth("D:D", decimal.NewFromInt(20))

	// SetFormula example
	sheet.AxisCell("F1").SetIntValue(100)
	sheet.AxisCell("F2").SetIntValue(200)
	sheet.AxisCell("F3").SetFormula("SUM(F1:F2)")

	fStyle := sheet.GetColStyle(ColumnNumber("F"))
	fStyle.Alignment.Horizontal = HorizontalAlignmentLeft
	fStyle.Alignment.Vertical = VerticalAlignmentCenter
	sheet.SetColStyle(ColumnNumber("F"), fStyle)

	err := f.SaveFile(docPath)
	assert.Nil(t, err)
}
