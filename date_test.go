package xlsx

import (
	"testing"
	"time"

	"github.com/stretchr/testify/assert"
)

func TestExcelTime(t *testing.T) {
	timeVal := time.Date(1980, 9, 8, 23, 40, 10, 40, time.UTC)
	excelTime := TimeToExcelTime(timeVal)
	assert.Equal(t, excelTime.String(), "29472.9862268518523149")
	assert.Equal(t, ExcelTimeToTime(excelTime), timeVal)
}
