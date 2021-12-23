## xlsReader

Библиотека для чтения xls файлов / Golang read xls library

## Установка / Installation

`$ go get github.com/shakinm/xlsReader`

## Использование / Using

```go
package main

import (
	"fmt"
	"github.com/shakinm/xlsReader/xls"
	"log"
)

func main() {

	workbook, err := xls.OpenFile("small_1_sheet.xls")

	if err!=nil {
		log.Panic(err.Error())
	}

	// Кол-во листов в книге
	// Number of sheets in the workbook
	//
	// for i := 0; i <= workbook.GetNumberSheets()-1; i++ {}

	fmt.Println(workbook.GetNumberSheets())

	sheet, err := workbook.GetSheet(0)

	if err!=nil {
		log.Panic(err.Error())
	}

	// Имя листа
	// Print sheet name
	println(sheet.GetName())

	// Вывести кол-во строк в листе
	// Print the number of rows in the sheet
	println(sheet.GetNumberRows())

	for i := 0; i <= sheet.GetNumberRows(); i++ {
		if row, err := sheet.GetRow(i); err == nil {
			if cell, err := row.GetCol(1); err == nil {

				// Значение ячейки, тип строка
				// Cell value, string type
				fmt.Println(cell.GetString())

				//fmt.Println(cell.GetInt64())
				//fmt.Println(cell.GetFloat64())

				// Тип ячейки (записи)
				// Cell type (records)
				fmt.Println(cell.GetType())

				// Получение отформатированной строки, например для ячеек с датой или проценты
				// Receiving a formatted string, for example, for cells with a date or a percentage
				xfIndex:=cell.GetXFIndex()
				formatIndex:=workbook.GetXFbyIndex(xfIndex)
				format:=workbook.GetFormatByIndex(formatIndex.GetFormatIndex())
				fmt.Println(format.GetFormatString(cell))

			}

		}
	}
}
```
 
