package xlsx

// Sheet sheet operator
type Sheet interface {
	// Name sheet name
	Name() string

	// SetCellValue set cell value
	//
	// Example:
	//     sheet.SetCellValue(1, 1, "val") // A1 => "val"
	//     sheet.SetCellValue(2, 3, 98.01) // B3 => 98.01
	//     sheet.SetCellValue(3, 1, 1000) // C1 => 1000
	//     sheet.SetCellValue(4, 4, time.Now()) // D4 => "2021-03-11 05:19"
	SetCellValue(col, row int, value any) Cell

	// SetAxisCellValue set cell value
	//
	// Example:
	//     sheet.SetAxisCellValue("A1", "val") // A1 => "val"
	//     sheet.SetAxisCellValue("B3", 98.01) // B3 => 98.01
	//     sheet.SetAxisCellValue("C1", 1000) // C1 => 1000
	//     sheet.SetAxisCellValue("D4", time.Now()) // D4 => "2021-03-11 05:19"
	SetAxisCellValue(axis Axis, value any) Cell

	// Cell get cell by cell col and row
	Cell(col, row int) Cell

	// AxisCell get cell by cell name
	AxisCell(axis Axis) Cell

	// SetColumnWidth set column width
	//
	// Example:
	//     sheet.SetColumnWidth("A:B", 20)
	SetColumnWidth(columnRange string, width int) Sheet

	// GetColumnWidth get column width
	//
	// Example:
	//	sheet.GetColumnWidth("A") // returns 20
	GetColumnWidth(columnName string) int

	// MergeCell merge cell
	//
	// Example:
	//     sheet.MergeCell("A1", "B1")
	MergeCell(start Axis, end Axis) Sheet

	// GetCellStyle get cell style
	GetCellStyle(col, row int) Style

	// GetAxisCellStyle get cell style
	GetAxisCellStyle(axis Axis) Style

	// SetCellStyle set cell style
	SetCellStyle(col, row int, style Style) Sheet

	// SetAxisCellStyle set cell style
	SetAxisCellStyle(axis Axis, style Style) Sheet

	// SetCellBorder set cell border
	//
	// Example:
	//     sheet.SetCellBorder(1, 1, BorderStyleThin, Color{Color: "FF0000"}, false, true, false, true)
	SetCellBorder(col, row int, borderStyle BorderStyle, borderColor Color, top, right, bottom, left bool) Sheet

	// SetAxisCellBorder set cell border
	//
	// Example:
	//     sheet.SetAxisCellBorder("A1", BorderStyleThin, Color{Color: "FF0000"}, false, true, false, true)
	SetAxisCellBorder(axis Axis, borderStyle BorderStyle, borderColor Color, top, right, bottom, left bool) Sheet
}
