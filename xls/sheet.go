package xls

import (
	"bytes"
	"errors"
	"github.com/shakinm/xlsReader/helpers"
	"github.com/shakinm/xlsReader/xls/record"
	"github.com/shakinm/xlsReader/xls/structure"
)

type rw struct {
	cols map[int]structure.CellData
}

type sheet struct {
	boundSheet    *record.BoundSheet
	rows          map[int]*rw
	wb            *Workbook
	maxCol        int // MaxCol index, countCol=MaxCol+1
	maxRow        int // MaxRow index, countRow=MaxRow+1
	hasAutofilter bool
}

func (s *sheet) GetName() (string) {
	return s.boundSheet.Rgch.String()
}

// Get row by index

func (s *sheet) GetRow(index int) (row *rw, err error) {

	if row, ok:= s.rows[index]; ok {
		return row , err
	} else {
		return row, errors.New("Out of range")
	}
}

func (rw *rw) GetCol(index int) (c structure.CellData, err error) {

	if col, ok:=rw.cols[index]; ok {
		return col, err
	} else {
		c = new(record.FakeBlank)
		return c, nil
	}



}

func (rw *rw) GetCols() (cols []structure.CellData) {

	for i := 0; i <= len(rw.cols)-1; i++ {
		if rw.cols[i] == nil {
			cols = append(cols, new(record.FakeBlank))
		} else {
			cols = append(cols, rw.cols[i])
		}
	}

	return cols
}

// Get all rows
func (s *sheet) GetRows() (rows []*rw) {

	for i := 0; i <= len(s.rows)-1; i++ {
		rows = append(rows, s.rows[i])
	}

	return rows
}

// Get number of rows
func (s *sheet) GetNumberRows() (n int) {
	return len(s.rows)
}

func (s *sheet) read(stream []byte) (err error) { // nolint: gocyclo

	var point int64
	point = int64(helpers.BytesToUint32(s.boundSheet.LbPlyPos[:]))
	var sPoint int64

	eof := false
Next:

	recordNumber := stream[point : point+2]
	recordDataLength := int64(helpers.BytesToUint16(stream[point+2 : point+4]))
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

	if bytes.Compare(recordNumber, []byte{0xFD, 0x00}) == 0 {
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
		c := new(record.BoolErr)
		c.Read(stream[sPoint : sPoint+recordDataLength])
		s.addCell(c, c.GetRow(), c.GetCol())
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
		rkRecords := c.GetArrayRKRecord()
		for i := 0; i <= len(rkRecords)-1; i++ {
			s.addCell(rkRecords[i].Get(), rkRecords[i].GetRow(), rkRecords[i].GetCol())
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
		delete(s.rows, 0)
	}

	return

}

func (s *sheet) addCell(cd structure.CellData, row [2]byte, column [2]byte) {

	r := int(helpers.BytesToUint16(row[:]))
	c := int(helpers.BytesToUint16(column[:]))

	if s.rows == nil {
		s.rows = map[int]*rw{}
	}
	if _, ok := s.rows[r]; !ok {
		s.rows[r] = new(rw)

		if _, ok := s.rows[r].cols[c]; !ok {

			colVal := map[int]structure.CellData{}
			colVal[c] = cd

			s.rows[r].cols = colVal
		}

	}

	s.rows[r].cols[c] = cd

}
