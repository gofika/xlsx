package xlsx

import (
	"math"
	"regexp"
	"strconv"
	"strings"
	"sync"
	"unicode/utf8"
)

var (
	colNames = make(map[byte]int)
	once     sync.Once
)

const (
	cols    = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
	colsLen = 26
)

func init() {
	once.Do(func() {
		for i := 0; i < colsLen; i++ {
			colNames[cols[i]] = i
		}
	})
}

// ColumnNumber convert the column name to column number
//
// Example:
//
//	xlsx.ColumnNumber("AY") // returns 51
func ColumnNumber(columnName string) (columnNumber int) {
	columnName = strings.ToUpper(columnName)

	if ok, _ := regexp.MatchString(`^[A-Z]{1,10}$`, columnName); !ok {
		return 0
	}
	digits := 0
	for i := len(columnName) - 1; i >= 0; i-- {
		c := columnName[i]
		columnNumber += (colNames[c] + 1) * int(math.Pow(colsLen, float64(digits)))
		digits++
	}
	return
}

// ColumnName convert the column number to column name
//
// Example:
//
//	xlsx.ColumnName(51) // returns "AY"
func ColumnName(columnNumber int) (columnName string) {
	if columnNumber < 1 {
		return ""
	}
	for {
		columnName = string([]byte{cols[0] + byte((columnNumber-1)%colsLen)}) + columnName
		columnNumber = (columnNumber - 1) / colsLen
		//columnNumber--
		if columnNumber <= 0 {
			break
		}
	}
	return
}

// CellNameToCoordinates convert cell name to [col, row] coordinates
//
// Example:
//
//	xlsx.CellNameToCoordinates("A1") // returns 1, 1
//	xlsx.CellNameToCoordinates("B5") // returns 2, 5
func CellNameToCoordinates(cellName string) (col int, row int) {
	cellName = strings.ToUpper(cellName)
	re := regexp.MustCompile(`^([A-Z]+)([1-9][0-9]*)$`)
	mcs := re.FindStringSubmatch(cellName)
	if len(mcs) < 3 {
		return 0, 0
	}
	col = ColumnNumber(mcs[1])
	row, _ = strconv.Atoi(mcs[2])
	return
}

// CoordinatesToCellName convert [col, row] coordinates to cell name
//
// Example:
//
//	xlsx.CoordinatesToCellName(1, 1) // returns "A1"
func CoordinatesToCellName(col, row int) string {
	if col < 1 || row < 1 {
		return ""
	}
	return ColumnName(col) + strconv.Itoa(row)
}

// filter and make sure sheet name valid
func trimSheetName(name string) string {
	const nameMaxLen = 31
	if strings.ContainsAny(name, `:\/?*[]`) || utf8.RuneCountInString(name) > nameMaxLen {
		r := make([]rune, 0, nameMaxLen)
		for _, v := range name {
			switch v {
			case 42, 47, 58, 63, 91, 92, 93: // replace :\/?*[]
				continue
			default:
				r = append(r, v)
			}
			if len(r) == nameMaxLen {
				break
			}
		}
		name = string(r)
	}
	return name
}
