package xls

import (
	"fmt"
	"testing"
)

func TestGetWorkBook(t *testing.T) {

	wb, err := OpenFile("./../testfie/small_1_sheet.xls")

	if err != nil {
		t.Error("Error: ", err)
	}

	s, err := wb.GetSheet(0)
	cells, _ := s.GetRow(2)

	for k := range cells.cols {
		c, _ := cells.GetCol(k)

		formatIndex := wb.GetXFbyIndex(c.GetXFIndex())
		format := wb.GetFormatByIndex(formatIndex.GetFormatIndex())

		if err == nil {
			switch k {
			case 0:
				if c.GetInt64() != 3 {
					t.Error("Expected 3, got ", c.GetInt64())
				}
			case 1:
				if c.GetInt64() != 4 {
					t.Error("Expected 4, got ", c.GetInt64())
				}
			case 2:
				if c.GetFloat64() != 2.1 {
					t.Error("Expected 2.1, got ", c.GetFloat64())
				}
			case 3:
				if c.GetString() != "String 3" {
					t.Error("Expected 'String 3', got ", c.GetString())
				}
			case 4:
				if c.GetString() != "https://github.com/shakinm/xlsReader" {
					t.Error("Expected 'https://github.com/shakinm/xlsReader', got ", c.GetString())
				}
			case 5:
				if c.GetString() != "" {
					t.Error("Expected '', got ", c.GetString())
				}
			case 6:
				if c.GetString() != "" {
					t.Error("Expected '', got ", c.GetString())
				}
			case 7:
				if c.GetString() != "String 3" {
					t.Error("Expected 'String 3', got ", c.GetString())
				}
			case 8:
				if c.GetInt64() != 3 {
					t.Error("Expected 3, got ", c.GetInt64())
				}
			case 9: // bool
				if c.GetInt64() != 1 {
					t.Error("Expected 1, got ", c.GetInt64())
				}
				if c.GetString() != "TRUE" {
					t.Error("Expected 'TRUE', got ", c.GetString())
				}
			case 10: //date
				if format.GetFormatString(c) != "9/3/19" {
					t.Error("Expected '9/3/19', got ", format.GetFormatString(c))
				}
			case 11: //dateTime
				if format.GetFormatString(c) != "09/03/2019 13:12:59" {
					t.Error("Expected '09/03/2019 13:12:59', got ", format.GetFormatString(c))
				}
			case 12:
				if format.GetFormatString(c) != "55.00%" {
					t.Error("Expected '55.00%', got ", format.GetFormatString(c))
				}
			case 13:
				if format.GetFormatString(c) != "#DIV/0!" {
					t.Error("Expected '#DIV/0!', got ", format.GetFormatString(c))
				}
			}

		}
	}
}

func TestMiniFatWorkBook(t *testing.T) {
	wb, err := OpenFile("./../testfie/table.xls")

	if err != nil {
		t.Error("Error: ", err)
	}

	for i := 0; i <= wb.GetNumberSheets()-1; i++ {

		sheet, _ := wb.GetSheet(i)
		if sheet.GetRows() != nil {
			for _, row := range sheet.GetRows() {
				if row  != nil {

					for _, col := range row.GetCols() {

					//	fmt.Println(col.GetString())
						xf := col.GetXFIndex()
						//fmt.Println(xf)
						style := wb.GetXFbyIndex(xf)
						//fmt.Println(style)
						formatIdx := style.GetFormatIndex()
						//fmt.Println(formatIdx)
						format := wb.GetFormatByIndex(formatIdx)
						//fmt.Println(format)

						fstr := format.GetFormatString(col)
						fmt.Println(fstr)

					}
				}

			}
		}

	}
}
