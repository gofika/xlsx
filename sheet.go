package xlsx

// Sheet sheet operator
type Sheet interface {
	// SetCellValue set cell value
	//
	// Example:
	//     sheet.SetCellValue(1, 1, "val") // A1 => "val"
	//     sheet.SetCellValue(2, 3, 98.01) // B3 => 98.01
	//     sheet.SetCellValue(3, 1, 1000) // C1 => 1000
	//     sheet.SetCellValue(4, 4, time.Now()) // D4 => "2021-03-11 05:19:16.483"
	SetCellValue(col, row int, value any) Cell

	// GetCellString get cell value of string
	//
	// Example:
	//     sheet.GetCellString(1, 1) // A1 => "val"
	GetCellString(col, row int) string

	// GetCellInt get cell value of string
	//
	// Example:
	//     sheet.GetCellInt(3, 1) // C1 => 1000
	GetCellInt(col, row int) int

	// Cell get cell by cell col and row
	Cell(col, row int) Cell

	// AxisCell get cell by cell name
	AxisCell(axis Axis) Cell
}
