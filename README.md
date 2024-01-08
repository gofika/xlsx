[![codecov](https://codecov.io/gh/gofika/xlsx/branch/main/graph/badge.svg)](https://codecov.io/gh/gofika/xlsx)
[![Build Status](https://github.com/gofika/xlsx/workflows/build/badge.svg)](https://github.com/gofika/xlsx)
[![go.dev](https://img.shields.io/badge/go.dev-reference-007d9c?logo=go&logoColor=white)](https://pkg.go.dev/github.com/gofika/xlsx)
[![Go Report Card](https://goreportcard.com/badge/github.com/gofika/xlsx)](https://goreportcard.com/report/github.com/gofika/xlsx)
[![Licenses](https://img.shields.io/github/license/gofika/xlsx)](LICENSE)
[![donate](https://img.shields.io/badge/Donate-PayPal-green.svg)](https://www.buymeacoffee.com/illi)

# xlsx

Microsoft .xlsx read/write for golang with high performance

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
    "time"

    "github.com/gofika/xlsx"
)

func main() {
    doc := xlsx.NewFile()

    // open default sheet "Sheet1"
    sheet := doc.OpenSheet("Sheet1")

    // write values
	valueCol := ColumnNumber("B")
    sheet.SetCellValue(xlsx.ColumnNumber("A"), 1, "Name") // A1 = Name
    sheet.SetCellValue(xlsx.ColumnNumber("A"), 2, "Jason") // A2 = Json
    sheet.SetCellValue(xlsx.ColumnNumber("B"), 1, "Score") // B1 = Score
    sheet.SetCellValue(xlsx.ColumnNumber("B"), 2, 100) // B2 = 100

    // time value
    sheet.SetAxisCellValue("C1", "Date") // C1 = Date
    sheet.SetAxisCellValue("C2", time.Date(1980, 9, 8, 23, 40, 10, 40, time.UTC)) // C2 = 1980-09-08 23:40

    // duration value
    sheet.SetAxisCellValue("D1", "Duration") // D1 = Duration
    sheet.SetAxisCellValue("D2", 30*time.Second) // D2 = 00:00:30

    // time value with custom format
    sheet.AxisCell("E1").SetStringValue("LastTime") // D1 = LastTime
    sheet.AxisCell("E2").
        SetTimeValue(time.Now()).
        SetNumberFormat("yyyy-mm-dd hh:mm:ss") // D2 = 2022-08-23 20:08:08 (your current time)

    // save to file
    if err := doc.SaveFile("Document1.xlsx"); err != nil {
        panic(err)
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
    // open exists document
    doc, err := xlsx.OpenFile("Document1.xlsx")
    if err != nil {
        panic(err)
        return
    }

    // open exists sheet
    sheet := doc.OpenSheet("Sheet2")

    // read cell string
    a1String := sheet.Cell(1, 1).GetStringValue()
    fmt.Println(a1String)

    // cell object read
    cell := sheet.AxisCell("B2")
    fmt.Println(cell.GetIntValue())
}
```

### Write spreadsheet as stream

Write document as a stream.

```go
package main

import (
    "io"
    "os"

    "github.com/gofika/xlsx"
)

func main() {
    // open file to write
    f, err := os.OpenFile("Document1.xlsx", os.O_CREATE|os.O_WRONLY, 0644)
    if err != nil {
        panic(err)
    }
    defer f.Close()

    doc := xlsx.NewFile()

    // do something with doc
    // ...

    // write to file or any io.Writer as stream
    doc.Save(f)
}
```

### NewFile with options

You can specify default configurations when calling **xlsx.NewFile**.

```go

package main

import (
    "io"
    "os"

    "github.com/gofika/xlsx"
)

func main() {
    // set document: default font name, default font size, default sheet name
    doc := xlsx.NewFile(xlsx.WithDefaultFontName("Arial"), xlsx.WithDefaultFontSize(12), xlsx.WithDefaultSheetName("Tab1"))

    // do something with doc
    // ...
}
```



## TODO:

- [x] Basic File Format
- [x] File: NewFile, OpenFile, SaveFile, Save, Sheets
- [x] Sheet:
    - [x] NewSheet, OpenSheet
    - [x] Name, SetCellValue, Cell, AxisCell, SetAxisCellValue, SetColumnWidth, GetColumnWidth, MergeCell
- [x] Cell:
    - [x] Row, Col
    - [x] SetValue, SetIntValue, SetFloatValue, SetFloatValuePrec, SetStringValue, SetBoolValue, SetDefaultValue, SetTimeValue, SetDateValue, SetDurationValue
    - [x] GetIntValue, GetStringValue, GetFloatValue, GetBoolValue, GetTimeValue, GetDurationValue
    - [x] SetNumberFormat
