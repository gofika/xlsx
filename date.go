package xlsx

import (
	"time"

	"github.com/shopspring/decimal"
)

const ()

var (
	since1900              = time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)
	secondsOfADay          = decimal.NewFromInt(int64((24 * time.Hour) / time.Second))
	nanosecondsOfADay      = decimal.NewFromInt(int64((24 * time.Hour) / time.Nanosecond))
	daysBetween1970And1900 = decimal.NewFromFloat(time.Date(1970, 1, 1, 0, 0, 0, 0, time.UTC).Sub(since1900).Hours() / 24)
)

// TimeToExcelTime convert time.Time to excel time format
func TimeToExcelTime(t time.Time) decimal.Decimal {
	return decimal.NewFromInt(t.Unix()).Div(secondsOfADay).Add(daysBetween1970And1900).Add(decimal.NewFromInt(int64(t.Nanosecond())).Div(nanosecondsOfADay))
}

// ExcelTimeToTime convert excel time format to time.Time
func ExcelTimeToTime(excelTime decimal.Decimal) time.Time {
	intPart := excelTime.IntPart()
	decimalPart := excelTime.Sub(decimal.NewFromInt(intPart))
	return since1900.Add(time.Duration(intPart) * time.Hour * 24).Add(time.Duration(decimalPart.Mul(nanosecondsOfADay).IntPart()) * time.Nanosecond)
}
