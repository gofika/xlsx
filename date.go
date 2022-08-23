package xlsx

import "time"

const (
	secondsOfADay     = float64((24 * time.Hour) / time.Second)
	nanosecondsOfADay = float64((24 * time.Hour) / time.Nanosecond)
)

var (
	daysBetween1970And1900 = float64(time.Date(1970, 1, 1, 0, 0, 0, 0, time.UTC).Sub(time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)) / (24 * time.Hour))
)

// TimeToExcelTime convert time.Time to excel time format
func TimeToExcelTime(t time.Time) float64 {
	return float64(t.Unix())/secondsOfADay + daysBetween1970And1900 + float64(t.Nanosecond())/nanosecondsOfADay
}
