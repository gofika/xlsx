package xlsx

import (
	"time"

	"github.com/shopspring/decimal"
)

// Cell cell operator
type Cell interface {
	// Row cell row number
	Row() int

	// Col cell col number
	Col() int

	// SetValue provides to set the value of a cell
	// Allow Types:
	//     int
	//     int8
	//     int16
	//     int32
	//     int64
	//     uint
	//     uint8
	//     uint16
	//     uint32
	//     uint64
	//     float32
	//     float64
	//     string
	//     bool
	//     time.Time
	//     time.Duration
	//     []byte
	//     decimal.Decimal
	//
	// Example:
	//     cell.SetValue(100)
	//     cell.SetValue("Hello")
	//     cell.SetValue(3.14)
	SetValue(value any) Cell

	// SetIntValue set cell for int type
	SetIntValue(value int) Cell

	// GetIntValue get cell value with int type
	GetIntValue() int

	// SetFloatValue set cell for decimal.Decimal type
	SetFloatValue(value decimal.Decimal) Cell

	// SetFloatValuePrec set cell for decimal.Decimal type with pres
	SetFloatValuePrec(value decimal.Decimal, prec int) Cell

	// GetFloatValue get cell value with decimal.Decimal type
	GetFloatValue() decimal.Decimal

	// SetStringValue set cell value for string type
	SetStringValue(value string) Cell

	// GetStringValue get cell value with string type
	GetStringValue() string

	// SetBoolValue set cell value for bool type
	SetBoolValue(value bool) Cell

	// GetBoolValue get cell value with bool type
	GetBoolValue() bool

	// SetDefaultValue set cell value without any type
	SetDefaultValue(value string) Cell

	// SetTimeValue set cell value for time.Time type
	SetTimeValue(value time.Time) Cell

	// GetTimeValue get cell value with time.Time type
	GetTimeValue() time.Time

	// SetDateValue set cell value for time.Time type with date format
	SetDateValue(value time.Time) Cell

	// SetDurationValue set cell value for time.Duration type
	SetDurationValue(value time.Duration) Cell

	// GetDurationValue get cell value with time.Duration type
	GetDurationValue() time.Duration

	// SetNumberFormat set cell number format with format code
	// https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.numberingformat?view=openxml-2.8.1
	SetNumberFormat(formatCode string) Cell
}
