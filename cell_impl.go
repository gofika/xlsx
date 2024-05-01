package xlsx

import (
	"fmt"
	"strconv"
	"time"

	"github.com/gofika/xlsx/packaging"
	"github.com/shopspring/decimal"
)

// cellImpl cell operator
type cellImpl struct {
	sheet    *sheetImpl
	col      int
	row      int
	cellName string
}

func newCell(sheet *sheetImpl, col, row int) *cellImpl {
	return &cellImpl{
		sheet:    sheet,
		col:      col,
		row:      row,
		cellName: CoordinatesToCellName(col, row),
	}
}

// Row cell row number
func (c *cellImpl) Row() int {
	return c.row
}

// Col cell col number
func (c *cellImpl) Col() int {
	return c.col
}

// Axis cell axis
func (c *cellImpl) Axis() Axis {
	return Axis(c.cellName)
}

func (c *cellImpl) getRow() *packaging.XRow {
	return c.sheet.getRow(c.row)
}

func (c *cellImpl) getCell() *packaging.XC {
	return c.sheet.getCell(c.col, c.row)
}

func (c *cellImpl) getSharedStrings() *sharedStrings {
	return c.sheet.file.ss
}

func (c *cellImpl) prepareCell() *packaging.XC {
	return c.sheet.prepareCell(c.col, c.row)
}

// func (c *cellImpl) prepareCellFormat() *packaging.XXf {
// 	return c.sheet.prepareCellFormat(c.col, c.row)
// }

func (c *cellImpl) GetStyle() Style {
	return c.sheet.GetCellStyle(c.col, c.row)
}

func (c *cellImpl) SetStyle(style Style) Cell {
	c.sheet.SetCellStyle(c.col, c.row, style)
	return c
}

func (c *cellImpl) SetValue(value any) Cell {
	switch v := value.(type) {
	case int, int8, int16, int32, int64, uint, uint8, uint16, uint32, uint64:
		c.setIntType(v)
	case float32:
		c.SetFloatValue(decimal.NewFromFloat32(v))
	case float64:
		c.SetFloatValue(decimal.NewFromFloat(v))
	case decimal.Decimal:
		c.SetFloatValue(v)
	case string:
		c.SetStringValue(v)
	case []byte:
		c.SetStringValue(string(v))
	case time.Duration:
		c.SetDurationValue(v)
	case time.Time:
		c.SetTimeValue(v)
	case bool:
		c.SetBoolValue(v)
	default:
		c.SetStringValue(fmt.Sprint(value))
	}
	return c
}

func (c *cellImpl) setIntType(value any) Cell {
	switch v := value.(type) {
	case int:
		c.SetIntValue(v)
	case int8:
		c.SetIntValue(int(v))
	case int16:
		c.SetIntValue(int(v))
	case int32:
		c.SetIntValue(int(v))
	case int64:
		c.SetIntValue(int(v))
	case uint:
		c.SetIntValue(int(v))
	case uint8:
		c.SetIntValue(int(v))
	case uint16:
		c.SetIntValue(int(v))
	case uint32:
		c.SetIntValue(int(v))
	case uint64:
		c.SetIntValue(int(v))
	}
	return c
}

// SetIntValue set cell for int type
func (c *cellImpl) SetIntValue(value int) Cell {
	cell := c.prepareCell()
	cell.T = ""
	cell.V = strconv.Itoa(value)
	return c
}

// GetIntValue get cell value with int type
func (c *cellImpl) GetIntValue() int {
	cell := c.getCell()
	if cell == nil {
		return 0
	}
	value, err := strconv.Atoi(cell.V)
	if err != nil {
		return 0
	}
	return value
}

// SetFloatValue set cell for decimal.Decimal type
func (c *cellImpl) SetFloatValue(value decimal.Decimal) Cell {
	c.SetFloatValuePrec(value, -1)
	return c
}

// GetFloatValue get cell value with decimal.Decimal type
func (c *cellImpl) GetFloatValue() decimal.Decimal {
	cell := c.getCell()
	if cell == nil {
		return decimal.Zero
	}
	value, err := decimal.NewFromString(cell.V)
	// value, err := strconv.ParseFloat(cell.V, 64)
	if err != nil {
		return decimal.Zero
	}
	return value
}

// SetFloatValuePrec set cell for decimal.Decimal type with pres
func (c *cellImpl) SetFloatValuePrec(value decimal.Decimal, prec int) Cell {
	cell := c.prepareCell()
	// cell.V = strconv.FormatFloat(value, 'f', prec, bitSize)
	if prec < 0 {
		cell.V = value.String()
	} else {
		cell.V = value.StringFixed(int32(prec))
	}
	return c
}

// GetStringValue get cell value with string type
func (c *cellImpl) GetStringValue() string {
	cell := c.getCell()
	if cell == nil {
		return ""
	}
	switch cell.T {
	case "str":
		return cell.V
	case "s":
		return c.getSharedStrings().Get(cell.V)
	default:
		return cell.V
	}
}

// SetStringValue set cell value for string type
func (c *cellImpl) SetStringValue(value string) Cell {
	cell := c.prepareCell()
	cell.T = "s"
	stringID := c.getSharedStrings().Append(value)
	cell.V = stringID
	return c
}

// SetBoolValue set cell value for bool type
func (c *cellImpl) SetBoolValue(value bool) Cell {
	cell := c.prepareCell()
	cell.T = "b"
	if value {
		cell.V = "1"
	} else {
		cell.V = "0"
	}
	return c
}

// GetBoolValue get cell value with bool type
func (c *cellImpl) GetBoolValue() bool {
	cell := c.getCell()
	if cell == nil {
		return false
	}
	return cell.V == "1"
}

// SetDefaultValue set cell value without any type
func (c *cellImpl) SetDefaultValue(value string) Cell {
	cell := c.prepareCell()
	cell.V = value
	return c
}

// SetTimeValue set cell value for time.Time type
func (c *cellImpl) SetTimeValue(value time.Time) Cell {
	cell := c.prepareCell()
	cell.T = ""

	excelTime := TimeToExcelTime(value)
	if excelTime.GreaterThan(decimal.Zero) {
		// cell.V = strconv.FormatFloat(excelTime, 'f', 5, 64)
		cell.V = excelTime.String()
		// cellFormat := c.prepareCellFormat()
		// cellFormat.ApplyNumberFormat = true
		// cellFormat.NumFmtID = 22 // built-in format. 22 = "yyyy-mm-dd hh:mm"

		style := c.GetStyle()
		style.NumberFormat.NumberFormatID = 22 // built-in format. 22 = "yyyy-mm-dd hh:mm"
		c.SetStyle(style)
	} else {
		cell.V = value.Format(time.RFC3339Nano)
	}
	return c
}

// GetTimeValue get cell value with time.Time type
func (c *cellImpl) GetTimeValue() time.Time {
	cell := c.getCell()
	if cell == nil {
		return time.Time{}
	}
	// excelTime, err := strconv.ParseFloat(cell.V, 64)
	excelTime, err := decimal.NewFromString(cell.V)
	if err != nil {
		return time.Time{}
	}
	return ExcelTimeToTime(excelTime)
}

// SetDateValue set cell value for time.Time type as date format
func (c *cellImpl) SetDateValue(value time.Time) Cell {
	cell := c.prepareCell()
	cell.T = ""

	excelTime := TimeToExcelTime(value)
	if excelTime.GreaterThan(decimal.Zero) {
		// cell.V = strconv.FormatFloat(excelTime, 'f', 5, 64)
		cell.V = excelTime.String()
		// cellFormat := c.prepareCellFormat()
		// cellFormat.ApplyNumberFormat = true
		// cellFormat.NumFmtID = 34 // built-in format. 34 = "yyyy-mm-dd"
		style := c.GetStyle()
		style.NumberFormat.NumberFormatID = 34 // built-in format. 34 = "yyyy-mm-dd"
		c.SetStyle(style)
	} else {
		cell.V = value.Format(time.RFC3339Nano)
	}
	return c
}

// SetDurationValue set cell value for time.Duration type
func (c *cellImpl) SetDurationValue(value time.Duration) Cell {
	cell := c.prepareCell()
	cell.V = strconv.FormatFloat(value.Seconds()/86400.0, 'f', 5, 64)
	// cellFormat := c.prepareCellFormat()
	// cellFormat.ApplyNumberFormat = true
	// cellFormat.NumFmtID = 21 // built-in format. 21 = "h:mm:ss"
	style := c.GetStyle()
	style.NumberFormat.NumberFormatID = 21 // built-in format. 21 = "h:mm:ss"
	c.SetStyle(style)
	return c
}

// GetDurationValue get cell value with time.Duration type
func (c *cellImpl) GetDurationValue() time.Duration {
	cell := c.getCell()
	if cell == nil {
		return 0
	}
	// excelTime, err := strconv.ParseFloat(cell.V, 64)
	excelTime, err := decimal.NewFromString(cell.V)
	if err != nil {
		return 0
	}
	// return time.Duration(excelTime * 86400.0 * time.Second)
	return time.Duration(excelTime.Mul(decimal.NewFromFloat(86400.0)).IntPart()) * time.Second
}

// SetFormula set cell formula
func (c *cellImpl) SetFormula(formula string) Cell {
	cell := c.prepareCell()
	cell.F = &packaging.XF{
		Content: formula,
	}
	return c
}

// GetFormula get cell formula
func (c *cellImpl) GetFormula() string {
	cell := c.getCell()
	if cell == nil || cell.F == nil {
		return ""
	}
	return cell.F.Content
}

// SetNumberFormat set cell number format with format code
// https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.numberingformat?view=openxml-2.8.1
func (c *cellImpl) SetNumberFormat(formatCode string) Cell {
	// numFmtID := c.sheet.prepareNumberingFormat(formatCode)
	// cellFormat := c.prepareCellFormat()
	// cellFormat.ApplyNumberFormat = true
	// cellFormat.NumFmtID = numFmtID
	style := c.GetStyle()
	style.NumberFormat.Format = formatCode
	style.NumberFormat.NumberFormatID = 0
	c.SetStyle(style)
	return c
}

func (c *cellImpl) SetCellBorder(borderStyle BorderStyle, borderColor Color, top, right, bottom, left bool) Cell {
	c.sheet.SetCellBorder(c.col, c.row, borderStyle, borderColor, top, right, bottom, left)
	return c
}
