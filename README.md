[![codecov](https://codecov.io/gh/gofika/xlsx/branch/main/graph/badge.svg)](https://codecov.io/gh/gofika/xlsx)
[![Build Status](https://github.com/gofika/xlsx/workflows/build/badge.svg)](https://github.com/gofika/xlsx)
[![go.dev](https://img.shields.io/badge/go.dev-reference-007d9c?logo=go&logoColor=white)](https://pkg.go.dev/github.com/gofika/xlsx)
[![Go Report Card](https://goreportcard.com/badge/github.com/gofika/xlsx)](https://goreportcard.com/report/github.com/gofika/xlsx)
[![Licenses](https://img.shields.io/github/license/gofika/xlsx)](LICENSE)
[![donate](https://img.shields.io/badge/Donate-PayPal-green.svg)](https://www.buymeacoffee.com/illi)

# xlsx

Microsoft .xlsx read/write for golang

## Basic Usage

### Installation

To get the package, execute:

```bash
go get github.com/gofika/xlsx
```

To import this package, add the following line to your code:

```js
import "github.com/gofika/xlsx"
```

### Create spreadsheet

Here is example usage that will create xlsx file.

```go
package main

import (
    "fmt"
    "github.com/gofika/xlsx"
    "time"
)

func main() {
    f := xlsx.NewFile()

    sheet := f.NewSheet("Sheet2")
    sheet.SetCellValue(xlsx.ColumnNumber("A"), 1, "Name")
    sheet.SetCellValue(xlsx.ColumnNumber("A"), 2, "Jason")
    sheet.SetCellValue(xlsx.ColumnNumber("B"), 1, "Score")
    sheet.SetCellValue(xlsx.ColumnNumber("B"), 2, 100)
    // date value
    sheet.SetCellValue(3, 1, "Date")
    sheet.Cell(3, 2).SetDateValue(time.Date(1980, 9, 8, 0, 0, 0, 0, time.Local))
    // time value
    sheet.AxisCell("D1").SetStringValue("LastTime")
    sheet.AxisCell("D2").
        SetTimeValue(time.Now()).
        SetNumberFormat("yyyy-mm-dd hh:mm:ss")

    if err := f.SaveFile("Document1.xlsx"); err != nil {
        fmt.Println(err)
    }
}
```

### Reading spreadsheet

The following constitutes the bare to read a spreadsheet document.

```go
package main

import (
    "fmt"
    "github.com/gofika/xlsx"
)

func main() {
    f, err := xlsx.OpenFile("Document1.xlsx")
    if err != nil {
        fmt.Println(err)
        return
    }

    sheet := f.OpenSheet("Sheet2")
    A1 := sheet.GetCellString(1, 1)
    fmt.Println(A1)

    cell := sheet.AxisCell("B2")
    fmt.Println(cell.GetIntValue())
}
```

## TODO:

- [x] Basic File Format
- [x] File: NewFile, OpenFile, SaveFile, Save, Sheets
- [ ] Sheet:
    - [x] NewSheet, OpenSheet
    - [x] SetCellValue, GetCellString, GetCellInt, Cell, AxisCell
    - [ ] ...
- [ ] Cell:
    - [x] Row, Col
    - [x] SetValue, SetIntValue, SetFloatValue, SetFloatValuePrec, SetStringValue, SetBoolValue, SetDefaultValue, SetTimeValue, SetDateValue, SetDurationValue
    - [x] GetIntValue, GetStringValue
    - [x] SetNumberFormat
    - [ ] ...
