package xlsx

import "time"

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
	//     []byte
	//     time.Duration
	//     time.Time
	//     bool
	//     nil
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

	// SetFloatValue set cell for float64 type
	SetFloatValue(value float64) Cell

	// SetFloatValuePrec set cell for float64 type with pres
	SetFloatValuePrec(value float64, prec int, bitSize int) Cell

	// GetStringValue get cell value with string type
	GetStringValue() string

	// SetStringValue set cell value for string type
	SetStringValue(value string) Cell

	// SetBoolValue set cell value for bool type
	SetBoolValue(value bool) Cell

	// SetDefaultValue set cell value without any type
	SetDefaultValue(value string) Cell

	// SetTimeValue set cell value for time.Time type
	SetTimeValue(value time.Time) Cell

	// SetDateValue set cell value for time.Time type with date format
	SetDateValue(value time.Time) Cell

	// SetDurationValue set cell value for time.Duration type
	SetDurationValue(value time.Duration) Cell

	// SetNumberFormat set cell number format with format code
	// https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.numberingformat?view=openxml-2.8.1
	SetNumberFormat(formatCode string) Cell
}
