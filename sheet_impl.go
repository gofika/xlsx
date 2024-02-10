package xlsx

import (
	"fmt"

	"github.com/gofika/xlsx/packaging"
	"github.com/shopspring/decimal"
)

// sheetImpl sheet operator
type sheetImpl struct {
	file       *fileImpl
	sheet      *packaging.XSheet
	sheetIndex int
}

func newSheet(file *fileImpl, sheet *packaging.XSheet, sheetIndex int) *sheetImpl {
	return &sheetImpl{
		file:       file,
		sheet:      sheet,
		sheetIndex: sheetIndex,
	}
}

func (s *sheetImpl) getWorksheet() *packaging.XWorksheet {
	return s.file.xFile.Worksheets[s.sheetIndex]
}

func (s *sheetImpl) getSheetData() *packaging.XSheetData {
	return s.getWorksheet().SheetData
}

func (s *sheetImpl) getStyleSheet() *packaging.XStyleSheet {
	return s.file.xFile.StyleSheet
}

// Name get sheet name
func (s *sheetImpl) Name() string {
	return s.sheet.Name
}

// SetCellValue set cell value
//
// Example:
//
//	sheet.SetCellValue(1, 1, "val") // A1 => "val"
//	sheet.SetCellValue(2, 3, 98.01) // B3 => 98.01
//	sheet.SetCellValue(3, 1, 1000) // C1 => 1000
//	sheet.SetCellValue(4, 4, time.Now()) // D4 => "2021-03-11 05:19:16.483"
func (s *sheetImpl) SetCellValue(col, row int, value any) (cell Cell) {
	cell = s.Cell(col, row).SetValue(value)
	return
}

// SetAxisCellValue set cell value
//
// Example:
//
//	sheet.SetAxisCellValue("A1", "val") // A1 => "val"
//	sheet.SetAxisCellValue("B3", 98.01) // B3 => 98.01
//	sheet.SetAxisCellValue("C1", 1000) // C1 => 1000
//	sheet.SetAxisCellValue("D4", time.Now()) // D4 => "2021-03-11 05:19"
func (s *sheetImpl) SetAxisCellValue(axis Axis, value any) (cell Cell) {
	cell = s.Cell(axis.C()).SetValue(value)
	return
}

// cell get cell by cell name
func (s *sheetImpl) Cell(col, row int) Cell {
	return newCell(s, col, row)
}

// AxisCell get cell by cell name
func (s *sheetImpl) AxisCell(axis Axis) Cell {
	return s.Cell(axis.C())
}

func (s *sheetImpl) getRow(row int) *packaging.XRow {
	sheetData := s.getSheetData()
	for _, r := range sheetData.Row {
		if r.R == row {
			return r
		}
	}
	return nil
}

func (s *sheetImpl) prepareRow(row int) *packaging.XRow {
	r := s.getRow(row)
	if r != nil {
		return r
	}
	// create new row
	sheetData := s.getSheetData()
	r = &packaging.XRow{
		R: row,
	}
	rowIndex := row - 1

	if len(sheetData.Row) <= rowIndex { // empty slice or after last element
		sheetData.Row = append(sheetData.Row, r)
	} else {
		sheetData.Row = append(sheetData.Row[:rowIndex+1], sheetData.Row[rowIndex:]...)
		sheetData.Row[rowIndex] = r
	}
	return r
}

func (s *sheetImpl) getCell(col, row int) *packaging.XC {
	r := s.getRow(row)
	if r == nil {
		return nil
	}
	cellName := CoordinatesToCellName(col, row)
	for _, cell := range r.C {
		if cell.R == cellName {
			return cell
		}
	}
	return nil
}

func (s *sheetImpl) prepareCell(col, row int) *packaging.XC {
	cell := s.getCell(col, row)
	if cell != nil {
		return cell
	}
	// create new cell
	cellName := CoordinatesToCellName(col, row)
	r := s.prepareRow(row)
	cell = &packaging.XC{
		R: cellName,
	}

	// insert cell to row
	inserted := false
	for i := len(r.C) - 1; i >= 0; i-- {
		cCol, _ := CellNameToCoordinates(r.C[i].R)
		if cCol < col {
			r.C = append(r.C[:i+1], append([]*packaging.XC{cell}, r.C[i+1:]...)...)
			inserted = true
			break
		}
	}
	if !inserted {
		r.C = append([]*packaging.XC{cell}, r.C...)
	}

	// calc spans
	maxCol := 1
	for _, c := range r.C {
		spanCol, _ := CellNameToCoordinates(c.R)
		if spanCol > maxCol {
			maxCol = spanCol
		}
	}
	r.Spans = fmt.Sprintf("1:%d", maxCol)

	// prepare cell style
	worksheet := s.getWorksheet()
	if cell.S == 0 && worksheet.Cols != nil { // cell style not set && has col defines
		for _, c := range worksheet.Cols.Col {
			if c.Min <= col && col <= c.Max {
				cell.S = c.Style
			}
		}
	}

	return cell
}

func (s *sheetImpl) removeRow(row int) {
	sheetData := s.getSheetData()
	for i, r := range sheetData.Row {
		if r.R == row {
			sheetData.Row = append(sheetData.Row[:i], sheetData.Row[i+1:]...)
			return
		}
	}
}

func (s *sheetImpl) getCellFormat(col, row int) *packaging.XXf {
	cell := s.getCell(col, row)
	if cell == nil {
		return s.getDefaultCellFormat()
	}
	styleSheet := s.getStyleSheet()
	if cell.S < len(styleSheet.CellXfs.Xf) {
		return styleSheet.CellXfs.Xf[cell.S]
	}
	return s.getDefaultCellFormat()
}

func (s *sheetImpl) getDefaultCellFormat() *packaging.XXf {
	styleSheet := s.getStyleSheet()
	return styleSheet.CellXfs.Xf[0]
}

func (s *sheetImpl) getCellFont(col, row int) *packaging.XStyleSheetFont {
	cellFormat := s.getCellFormat(col, row)
	if cellFormat != nil {
		styleSheet := s.getStyleSheet()
		if cellFormat.FontID < len(styleSheet.Fonts.Font) {
			return styleSheet.Fonts.Font[cellFormat.FontID]
		}
	}
	return s.getDefaultFont()
}

func (s *sheetImpl) getDefaultFont() *packaging.XStyleSheetFont {
	styleSheet := s.getStyleSheet()
	return styleSheet.Fonts.Font[0]
}

func (s *sheetImpl) getCellAlignment(col, row int) *packaging.XAlignment {
	cellFormat := s.getCellFormat(col, row)
	return cellFormat.Alignment
}

func (s *sheetImpl) getCellBorder(col, row int) *packaging.XBorder {
	cellFormat := s.getCellFormat(col, row)
	styleSheet := s.getStyleSheet()
	if cellFormat.BorderID < len(styleSheet.Borders.Border) {
		return styleSheet.Borders.Border[cellFormat.BorderID]
	}
	return s.getDefaultBorder()
}

func (s *sheetImpl) getDefaultBorder() *packaging.XBorder {
	styleSheet := s.getStyleSheet()
	return styleSheet.Borders.Border[0]
}

func (s *sheetImpl) getCellFill(col, row int) *packaging.XFill {
	cellFormat := s.getCellFormat(col, row)
	styleSheet := s.getStyleSheet()
	if cellFormat.FillID < len(styleSheet.Fills.Fill) {
		return styleSheet.Fills.Fill[cellFormat.FillID]
	}
	return s.getDefaultFill()
}

func (s *sheetImpl) getDefaultFill() *packaging.XFill {
	styleSheet := s.getStyleSheet()
	return styleSheet.Fills.Fill[0]
}

// func (s *sheetImpl) prepareCellFormat(col, row int) *packaging.XXf {
// 	cell := s.prepareCell(col, row)
// 	var cellFormat *packaging.XXf
// 	if cell.S > 0 { // 0=default cant modify
// 		cellFormat = s.getCellFormat(col, row)
// 		return cellFormat
// 	}
// 	// create new xf
// 	cellFormat = &packaging.XXf{
// 		NumFmtID: 0,
// 		FontID:   0,
// 		FillID:   0,
// 		BorderID: 0,
// 		XfID:     0,
// 	}
// 	styleSheet := s.getStyleSheet()
// 	styleSheet.CellXfs.Xf = append(styleSheet.CellXfs.Xf, cellFormat)
// 	styleSheet.CellXfs.Count = len(styleSheet.CellXfs.Xf)
// 	cell.S = styleSheet.CellXfs.Count - 1
// 	return cellFormat
// }

func (s *sheetImpl) GetCellStyle(col, row int) Style {
	cellFormat := s.getCellFormat(col, row)
	var style Style
	// Font
	style.Font = fontFromPackaing(s.getCellFont(col, row))

	// Alignment
	style.Alignment = alignmentFromPackaging(s.getCellAlignment(col, row))

	// Border
	style.Border = borderFromPackaging(s.getCellBorder(col, row))

	// Fill
	style.Fill = fillFromPackaging(s.getCellFill(col, row))

	// IncludeQuotePrefix
	style.IncludeQuotePrefix = cellFormat.QuotePrefix.Value()

	// NumberFormat
	style.NumberFormat = numberFormatFromPackaging(s.findNumFmt(cellFormat.NumFmtID))
	return style
}

func (s *sheetImpl) GetAxisCellStyle(axis Axis) Style {
	return s.GetCellStyle(axis.C())
}

func (s *sheetImpl) SetCellStyle(col, row int, style Style) Sheet {
	styleSheet := s.getStyleSheet()
	var format packaging.XXf
	// Font
	fontID := -1
	font := packaingToFont(style.Font)
	for i, f := range styleSheet.Fonts.Font {
		if font.Equal(f) {
			fontID = i
			break
		}
	}
	if fontID == -1 {
		fontID = len(styleSheet.Fonts.Font)
		styleSheet.Fonts.Font = append(styleSheet.Fonts.Font, font)
		styleSheet.Fonts.Count = len(styleSheet.Fonts.Font)
	}
	format.FontID = fontID
	format.ApplyFont = packaging.NewBool(fontID > 0)

	// Alignment
	alignment := packaingToAlignment(style.Alignment)
	if !alignment.IsZero() {
		format.Alignment = alignment
	}
	format.ApplyAlignment = packaging.NewBool(!alignment.IsZero())

	// Border
	border := packaingToBorder(style.Border)
	borderID := -1
	for i, b := range styleSheet.Borders.Border {
		if border.Equal(b) {
			borderID = i
			break
		}
	}
	if borderID == -1 {
		borderID = len(styleSheet.Borders.Border)
		styleSheet.Borders.Border = append(styleSheet.Borders.Border, border)
		styleSheet.Borders.Count = len(styleSheet.Borders.Border)
	}
	format.BorderID = borderID
	format.ApplyBorder = packaging.NewBool(borderID > 0)

	// Fill
	fill := packaingToFill(style.Fill)
	fillID := -1
	for i, f := range styleSheet.Fills.Fill {
		if fill.Equal(f) {
			fillID = i
			break
		}
	}
	if fillID == -1 {
		fillID = len(styleSheet.Fills.Fill)
		styleSheet.Fills.Fill = append(styleSheet.Fills.Fill, fill)
		styleSheet.Fills.Count = len(styleSheet.Fills.Fill)
	}
	format.FillID = fillID
	format.ApplyFill = packaging.NewBool(fillID > 0)

	// IncludeQuotePrefix
	format.QuotePrefix = packaging.NewBool(style.IncludeQuotePrefix)

	// NumberFormat
	numFmt := packaingToNumberFormat(style.NumberFormat)
	numFmtID := -1
	if numFmt.NumFmtID > 0 && numFmt.NumFmtID < BuiltInNumFmtMax {
		// built-in format
		numFmtID = numFmt.NumFmtID
	} else {
		if numFmt.NumFmtID == 0 && numFmt.FormatCode != "" {
			for i, formatCode := range builtInNumFmt {
				if numFmt.FormatCode == formatCode {
					numFmtID = i
					break
				}
			}
			if numFmtID == -1 {
				// custom format
				for i, nf := range styleSheet.NumFmts.NumFmt {
					if numFmt.FormatCode == nf.FormatCode {
						numFmtID = i
						break
					}
				}
			}
		} else {
			for i, nf := range styleSheet.NumFmts.NumFmt {
				if numFmt.Equal(nf) {
					numFmtID = i
					break
				}
			}
		}
	}
	if numFmtID == -1 {
		if numFmt.FormatCode != "" {
			numFmtID = len(styleSheet.NumFmts.NumFmt) + BuiltInNumFmtMax + 1
			numFmt.NumFmtID = numFmtID
			styleSheet.NumFmts.NumFmt = append(styleSheet.NumFmts.NumFmt, numFmt)
			styleSheet.NumFmts.Count = len(styleSheet.NumFmts.NumFmt)
		} else {
			numFmtID = 0
		}
	}
	format.NumFmtID = numFmtID
	format.ApplyNumberFormat = packaging.NewBool(numFmtID > 0)

	// set cell format
	formatID := -1
	for i, f := range styleSheet.CellXfs.Xf {
		if f.Equal(&format) {
			formatID = i
			break
		}
	}
	cell := s.prepareCell(col, row)
	if formatID == -1 {
		formatID = len(styleSheet.CellXfs.Xf)
		styleSheet.CellXfs.Xf = append(styleSheet.CellXfs.Xf, &format)
		styleSheet.CellXfs.Count = len(styleSheet.CellXfs.Xf)
	}
	cell.S = formatID
	return s
}

func (s *sheetImpl) SetAxisCellStyle(axis Axis, style Style) Sheet {
	col, row := axis.C()
	return s.SetCellStyle(col, row, style)
}

func (s *sheetImpl) findNumFmt(numFmtID int) (numFmt *packaging.XNumFmt) {
	numFmts := s.prepareNumFmts()
	for _, nf := range numFmts.NumFmt {
		if nf.NumFmtID == numFmtID {
			return nf
		}
	}
	nf, ok := builtInNumFmt[numFmtID]
	if ok {
		return &packaging.XNumFmt{
			FormatCode: nf,
			NumFmtID:   numFmtID,
		}
	}
	return nil
}

func (s *sheetImpl) findNumFmtCode(formatCode string) (numFmt *packaging.XNumFmt) {
	numFmts := s.prepareNumFmts()
	for _, nf := range numFmts.NumFmt {
		if nf.FormatCode == formatCode {
			return nf
		}
	}
	// try find built in format
	for i, nf := range builtInNumFmt {
		if nf == formatCode {
			return &packaging.XNumFmt{
				FormatCode: formatCode,
				NumFmtID:   i,
			}
		}
	}
	return nil
}

func (s *sheetImpl) prepareNumFmts() *packaging.XNumFmts {
	styleSheet := s.getStyleSheet()
	if styleSheet.NumFmts == nil {
		styleSheet.NumFmts = &packaging.XNumFmts{
			Count:  0,
			NumFmt: []*packaging.XNumFmt{},
		}
	}
	return styleSheet.NumFmts
}

// SetColumnWidth set column width
//
// Example:
//
//	sheet.SetColumnWidth("A:B", 20)
func (s *sheetImpl) SetColumnWidth(columnRange string, width int) Sheet {
	min, max := ColumnRange(columnRange)
	if min == 0 || max == 0 {
		return s
	}
	worksheet := s.getWorksheet()
	if worksheet.Cols == nil {
		worksheet.Cols = &packaging.XCols{
			Col: []*packaging.XCol{},
		}
	}
	for _, c := range worksheet.Cols.Col {
		if c.Min == min && c.Max == max {
			c.Width = decimal.NewFromInt32(int32(width))
			return s
		}
	}
	worksheet.Cols.Col = append(worksheet.Cols.Col, &packaging.XCol{
		Min:   min,
		Max:   max,
		Width: decimal.NewFromInt32(int32(width)),
	})
	return s
}

// GetColumnWidth get column width
//
// Example:
//
//	sheet.GetColumnWidth("A") // returns 20
func (s *sheetImpl) GetColumnWidth(columnName string) int {
	col := ColumnNumber(columnName)
	worksheet := s.getWorksheet()
	if worksheet.Cols != nil {
		for _, c := range worksheet.Cols.Col {
			if c.Min <= col && col <= c.Max {
				return int(c.Width.IntPart())
			}
		}
	}
	return 0
}

// MergeCell merge cell
//
// Example:
//
//	sheet.MergeCell("A1", "B1")
func (s *sheetImpl) MergeCell(start Axis, end Axis) Sheet {
	startCol, startRow := start.C()
	endCol, endRow := end.C()
	if startCol == endCol && startRow == endRow {
		return s
	}
	// convert B1:A2 to A1:B2
	if endCol < startCol {
		startCol, endCol = endCol, startCol
	}
	if endRow < startRow {
		startRow, endRow = endRow, startRow
	}
	sheetData := s.getSheetData()
	// remove cell
	for row := startRow; row <= endRow; row++ {
		for col := startCol; col <= endCol; col++ {
			cell := s.getCell(col, row)
			if cell != nil {
				for i, c := range sheetData.Row[row-1].C {
					if c.R == cell.R {
						sheetData.Row[row-1].C = append(sheetData.Row[row-1].C[:i], sheetData.Row[row-1].C[i+1:]...)
						break
					}
				}
			}
		}
	}
	// merge cell
	startCellName := CoordinatesToCellName(startCol, startRow)
	endCellName := CoordinatesToCellName(endCol, endRow)
	mergeCell := &packaging.XMergeCell{
		Ref: fmt.Sprintf("%s:%s", startCellName, endCellName),
	}
	worksheet := s.getWorksheet()
	if worksheet.MergeCells == nil {
		worksheet.MergeCells = &packaging.XMergeCells{
			Count:     0,
			MergeCell: []*packaging.XMergeCell{},
		}
	}
	// check merge cell exist
	for _, mc := range worksheet.MergeCells.MergeCell {
		if mc.Ref == mergeCell.Ref {
			return s
		}
	}
	worksheet.MergeCells.MergeCell = append(worksheet.MergeCells.MergeCell, mergeCell)
	worksheet.MergeCells.Count = len(worksheet.MergeCells.MergeCell)
	return s
}

func (s *sheetImpl) SetCellBorder(col, row int, borderStyle BorderStyle, borderColor Color, top, right, bottom, left bool) Sheet {
	style := s.GetCellStyle(col, row)
	backgroundColor := style.Fill.BackgroundColor
	if backgroundColor.IsZero() {
		backgroundColor = Color{Color: "FFFFFFFF"}
	}
	style.Border.TopBorder = borderStyle
	if top {
		style.Border.TopBorderColor = borderColor
	} else {
		style.Border.TopBorderColor = backgroundColor
	}
	style.Border.BottomBorder = borderStyle
	if bottom {
		style.Border.BottomBorderColor = borderColor
	} else {
		style.Border.BottomBorderColor = backgroundColor
	}
	style.Border.LeftBorder = borderStyle
	if left {
		style.Border.LeftBorderColor = borderColor
	} else {
		style.Border.LeftBorderColor = backgroundColor
	}
	style.Border.RightBorder = borderStyle
	if right {
		style.Border.RightBorderColor = borderColor
	} else {
		style.Border.RightBorderColor = backgroundColor
	}
	s.SetCellStyle(col, row, style)
	return s
}

func (s *sheetImpl) SetAxisCellBorder(axis Axis, borderStyle BorderStyle, borderColor Color, top, right, bottom, left bool) Sheet {
	col, row := axis.C()
	return s.SetCellBorder(col, row, borderStyle, borderColor, top, right, bottom, left)
}
