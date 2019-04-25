package xls

import (
	"fmt"
	"testing"
)

func TestGetWorkBook(t *testing.T) {
	wb, err := OpenFile("./../testfie/small_1_sheet.xls")

	if err != nil {
		fmt.Println("sdf")
	}
	s, err := wb.GetSheet(0)

	//fmt.Println(s.GetName())

	cells, _ := s.GetRow(2)
	//c, _ := cells.GetCol(12)
	//fmt.Println(c.String())
	for k := range cells.cols {
		c, _ := cells.GetCol(k)
		xfIndex:=c.GetXFIndex()
		formatIndex:=wb.xf[xfIndex].GetFormatIndex()
		format:=wb.formats[formatIndex]


		if err == nil {
			fmt.Println(format.GetFormatString()  )
			fmt.Println(c.GetType())
			fmt.Println(format.GetIndex())
			fmt.Println(c.GetString())
			fmt.Println("--------------------------------------------")

		}
	}

}
