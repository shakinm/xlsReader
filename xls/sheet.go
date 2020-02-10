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

type Sheet struct {
	boundSheet    *record.BoundSheet
	rows          map[int]*rw
	wb            *Workbook
	maxCol        int // maxCol index, countCol=maxCol+1
	maxRow        int // maxRow index, countRow=maxRow+1
	hasAutofilter bool
}

func (s *Sheet) GetName() string {
	return s.boundSheet.GetName()
}

// Get row by index

func (s *Sheet) GetRow(index int) (row *rw, err error) {

	if row, ok := s.rows[index]; ok {
		return row, err
	} else {
		return row, errors.New("Out of range")
	}
}

func (rw *rw) GetCol(index int) (c structure.CellData, err error) {

	if col, ok := rw.cols[index]; ok {
		return col, err
	} else {
		c = new(record.FakeBlank)
		return c, nil
	}

}

func (rw *rw) GetCols() (cols []structure.CellData) {

	var maxColKey int

	for k, _ := range rw.cols {
		if k > maxColKey {
			maxColKey = k
		}
	}

	for i := 0; i <= maxColKey; i++ {
		if rw.cols[i] == nil {
			cols = append(cols, new(record.FakeBlank))
		} else {
			cols = append(cols, rw.cols[i])
		}
	}

	return cols
}

// Get all rows
func (s *Sheet) GetRows() (rows []*rw) {
	for _, v := range s.rows {
		rows = append(rows, v)
	}

	return rows
}

// Get number of rows
func (s *Sheet) GetNumberRows() (n int) {
	return len(s.rows)
}

func (s *Sheet) read(stream []byte) (err error) { // nolint: gocyclo

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

	//LABEL - Cell Value, String Constant
	if bytes.Compare(recordNumber, record.LabelRecord[:]) == 0 {
		if bytes.Compare(s.wb.vers[:], record.FlagBIFF8) == 0 {
			c := new(record.LabelBIFF8)
			c.Read(stream[sPoint : sPoint+recordDataLength])
			s.addCell(c, c.GetRow(), c.GetCol())
		} else {
			c := new(record.LabelBIFF5)
			c.Read(stream[sPoint : sPoint+recordDataLength])
			s.addCell(c, c.GetRow(), c.GetCol())
		}

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
		blRecords := c.GetArrayBlRecord()
		for i := 0; i <= len(blRecords)-1; i++ {
			s.addCell(blRecords[i].Get(), blRecords[i].GetRow(), blRecords[i].GetCol())
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

func (s *Sheet) addCell(cd structure.CellData, row [2]byte, column [2]byte) {

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
