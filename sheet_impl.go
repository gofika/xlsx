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
		return nil
	}
	styleSheet := s.getStyleSheet()
	if cell.S < len(styleSheet.CellXfs.Xf) {
		return styleSheet.CellXfs.Xf[cell.S]
	}
	return nil
}

func (s *sheetImpl) prepareCellFormat(col, row int) *packaging.XXf {
	cell := s.prepareCell(col, row)
	var cellFormat *packaging.XXf
	if cell.S > 0 { // 0=default cant modify
		cellFormat = s.getCellFormat(col, row)
		if cellFormat != nil {
			return cellFormat
		}
	}
	// create new xf
	cellFormat = &packaging.XXf{
		NumFmtID: 0,
		FontID:   0,
		FillID:   0,
		BorderID: 0,
		XfID:     0,
	}
	styleSheet := s.getStyleSheet()
	styleSheet.CellXfs.Xf = append(styleSheet.CellXfs.Xf, cellFormat)
	styleSheet.CellXfs.Count = len(styleSheet.CellXfs.Xf)
	cell.S = styleSheet.CellXfs.Count - 1
	return cellFormat
}

func (s *sheetImpl) prepareNumberingFormat(formatCode string) (numFmtID int) {
	styleSheet := s.getStyleSheet()
	if styleSheet.NumFmts == nil {
		styleSheet.NumFmts = &packaging.XNumFmts{
			Count:  0,
			NumFmt: []*packaging.XNumFmt{},
		}
	}
	for _, numFmt := range styleSheet.NumFmts.NumFmt {
		if numFmt.FormatCode == formatCode {
			return numFmt.NumFmtId
		}
	}
	// create new numFmt
	numFmt := &packaging.XNumFmt{
		FormatCode: formatCode,
		NumFmtId:   1000 + len(styleSheet.NumFmts.NumFmt),
	}
	styleSheet.NumFmts.NumFmt = append(styleSheet.NumFmts.NumFmt, numFmt)
	styleSheet.NumFmts.Count = len(styleSheet.NumFmts.NumFmt)
	return numFmt.NumFmtId
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
