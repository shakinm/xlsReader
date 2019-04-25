package xls

import (
	"bytes"
	"errors"
	"github.com/shakinm/xlsReader/helpers"
	"github.com/shakinm/xlsReader/xls/record"
	"github.com/shakinm/xlsReader/xls/structure"
	"github.com/jinzhu/copier"
)

type rw struct {
	cols []col
}

type col struct {
	cell structure.CellData
}



type sheet struct {
	boundSheet    *record.BoundSheet
	rows          []rw
	wb            *Workbook
	maxCol        int // MaxCol index, countCol=MaxCol+1
	maxRow        int // MaxRow index, countRow=MaxRow+1
	hasAutofilter bool
}

func (s *sheet) GetName() (string) {
		return s.boundSheet.Rgch.String()
}

// Get row by index

func (s *sheet) GetRow(index int) (row rw, err error) {

	if len(s.rows)-1 < index {
		return row, errors.New("Out of range")
	}
	return s.rows[index], nil
}

func (rw *rw) GetCol(index int) (c structure.CellData, err error) {
	if len(rw.cols)-1 < index {
		return c, errors.New("Out of range")
	}

	if rw.cols[index].cell==nil {
		c=new(record.FakeBlank)
		return c, nil
	}
	return rw.cols[index].cell, nil
}

// Get all rows
func (s *sheet) GetRows(index int) (rows []rw, err error) {
	return s.rows, nil
}

// Get number of rows
func (s *sheet) GetNumberRows() (n int) {
	return s.maxRow + 1
}

func (s *sheet) read(stream []byte) (err error) { // nolint: gocyclo

	var point int32
	point = int32(helpers.BytesToUint32(s.boundSheet.LbPlyPos[:]))
	var sPoint int32

	eof := false
Next:

	recordNumber := stream[point : point+2]
	recordDataLength := int32(helpers.BytesToUint16(stream[point+2 : point+4]))
	sPoint = point + 4


	if bytes.Compare(recordNumber, record.AutofilterInfoRecord[:]) == 0 {
		c := new(record.AutofilterInfo)
		c.Read(stream[sPoint : sPoint+recordDataLength])
		if c.GetCountEntries() > 0 {
			s.hasAutofilter = true
		} else {
			s.hasAutofilter = false
		}
		goto EIF

	}

	//LABELSST - String constant that uses BIFF8 shared string table (new to BIFF8)
	if bytes.Compare(recordNumber, record.LabelSStRecord[:]) == 0 {
		c := new(record.LabelSSt)
		c.Read(stream[sPoint:sPoint+recordDataLength], &s.wb.sst)
		s.addCell(c, c.GetRow(), c.GetCol())
		goto EIF
	}

	if bytes.Compare(recordNumber, []byte{0xFD,0x00}) == 0 {
		//todo: сделать
		goto EIF
	}

	//ARRAY - An array-entered formula
	if bytes.Compare(recordNumber, record.ArrayRecord[:]) == 0 {
		//todo: сделать
		goto EIF
	}
	//BLANK - An empty col
	if bytes.Compare(recordNumber, record.BlankRecord[:]) == 0 {
		c := new(record.Blank)
		c.Read(stream[sPoint : sPoint+recordDataLength])
		s.addCell(c, c.GetRow(), c.GetCol())
		goto EIF
	}

	//BOOLERR - A Boolean or error value
	if bytes.Compare(recordNumber, record.BoolErrRecord[:]) == 0 {
		//todo: сделать
		goto EIF
	}

	//FORMULA - A col formula, stored as parse tokens
	if bytes.Compare(recordNumber, record.FormulaRecord[:]) == 0 {
		//todo: сделать
		goto EIF
	}

	//NUMBER  - An IEEE floating-point number
	if bytes.Compare(recordNumber, record.NumberRecord[:]) == 0 {
		c := new(record.Number)
		c.Read(stream[sPoint : sPoint+recordDataLength])
		s.addCell(c, c.GetRow(), c.GetCol())
		goto EIF
	}

	//MULBLANK - Multiple empty rows (new to BIFF5)
	if bytes.Compare(recordNumber, record.MulBlankRecord[:]) == 0 {
		c := new(record.MulBlank)
		c.Read(stream[sPoint : sPoint+recordDataLength])
		for _, bl := range c.GetArrayBlRecord() {
			s.addCell(&bl, bl.GetRow(), bl.GetCol())
		}
		goto EIF
	}

	//RK - An RK number
	if bytes.Compare(recordNumber, record.RkRecord[:]) == 0 {
		c := new(record.Rk)
		c.Read(stream[sPoint : sPoint+recordDataLength])
		s.addCell(c, c.GetRow(), c.GetCol())
		goto EIF
	}

	//MULRK - Multiple RK numbers (new to BIFF5)
	if bytes.Compare(recordNumber, record.MulRKRecord[:]) == 0 {
		c := new(record.MulRk)
		c.Read(stream[sPoint : sPoint+recordDataLength])
		for _, rk := range c.GetArrayRKRecord() {
			 cl := record.Rk{}
			copier.Copy(&cl , &rk)
			s.addCell(cl.Get(), cl.GetRow(), cl.GetCol())
		}
		goto EIF

	}

	//RSTRING - Cell with character formatting
	if bytes.Compare(recordNumber, record.RStringRecord[:]) == 0 {
		//todo: сделать
		goto EIF
	}

	//SHRFMLA - A shared formula (new to BIFF5)
	if bytes.Compare(recordNumber, record.SharedFormulaRecord[:]) == 0 {
		//todo: сделать
		goto EIF
	}

	//STRING - A string that represents the result of a formula
	if bytes.Compare(recordNumber, record.StringRecord[:]) == 0 {
		//todo: сделать
		goto EIF
	}

	if bytes.Compare(recordNumber, record.RowRecord[:]) == 0 {
		//todo: сделать
		goto EIF
	}

	//EOF
	if bytes.Compare(recordNumber, record.EOFRecord[:]) == 0 && recordDataLength == 0 {
		eof = true
	}
EIF:
	point = point + recordDataLength + 4
	if !eof {
		goto Next
	}

	// Trim empty  row and skip 0 rows with autofilters
	if s.hasAutofilter {
		s.maxRow = s.maxRow - 1
		s.rows = s.rows[1:s.maxRow]
	}
	s.rows = s.rows[:s.maxRow+1]

	return

}

func (s *sheet) addCell(cd structure.CellData, row [2]byte, column [2]byte) {

	r := int(helpers.BytesToUint16(row[:]))

	c := int(helpers.BytesToUint16(column[:]))

	if s.maxCol < c {
		s.maxCol = c
	}
	if s.maxRow < r {
		s.maxRow = r
	}

	if len(s.rows) < r+1 {
		newRows := make([]rw, r+2000)
		copy(newRows, s.rows)
		s.rows = newRows
	}

	if len(s.rows[r].cols) < c+1 {
		newCols := make([]col, c+1)
		copy(newCols, s.rows[r].cols)
		s.rows[r].cols = newCols
	}


	s.rows[r].cols[c].cell = cd

}
