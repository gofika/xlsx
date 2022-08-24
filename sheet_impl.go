package xlsx

import (
	"fmt"

	"github.com/gofika/xlsx/packaging"
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

// GetCellString get cell value of string
//
// Example:
//
//	sheet.GetCellString(1, 1) // A1 => "val"
func (s *sheetImpl) GetCellString(col, row int) string {
	return s.Cell(col, row).GetStringValue()
}

// GetCellInt get cell value of string
//
// Example:
//
//	sheet.GetCellInt(3, 1) // C1 => 1000
func (s *sheetImpl) GetCellInt(col, row int) int {
	return s.Cell(col, row).GetIntValue()
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
	r.C = append(r.C, cell)
	if len(r.C) > 1 {
		r.Spans = fmt.Sprintf("1:%d", len(r.C))
	}

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
