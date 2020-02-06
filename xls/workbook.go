package xls

import (
	"bytes"
	"errors"
	"github.com/shakinm/xlsReader/helpers"
	"github.com/shakinm/xlsReader/xls/record"
)

// Workbook struct
type Workbook struct {
	sheets   []Sheet
	codepage record.CodePage
	sst      record.SST
	xf       []record.XF
	formats  map[int]record.Format
	vers     [2]byte
}

// GetNumberSheets - Number of sheets in the workbook
func (wb *Workbook) GetNumberSheets() int {
	return len(wb.sheets)
}

// GetSheets - Get sheets in the workbook
func (wb *Workbook) GetSheets() []Sheet {
	return wb.sheets
}

// GetSheet - Get Sheet by ID
func (wb *Workbook) GetSheet(sheetID int) (sheet *Sheet, err error) { // nolint: golint

	if len(wb.sheets) >= 1 && len(wb.sheets) >= sheetID {
		return &wb.sheets[sheetID], err
	}

	return nil, errors.New("error. Sheet not found")
}

// GetXF -  Return Extended Format Record by index
func (wb *Workbook) GetXFbyIndex(index int) record.XF {
	if len(wb.xf)<index {
		return wb.xf[15]
	}
	return wb.xf[index]
}

// GetXF -  Return FORMAT record describes a number format in the workbook
func (wb *Workbook) GetFormatByIndex(index int) record.Format {
	return wb.formats[index]
}

// GetCodePage - codepage
func (wb *Workbook) GetCodePage() record.CodePage {
	return wb.codepage
}

func (wb *Workbook) addSheet(bs *record.BoundSheet) (sheet Sheet) { // nolint: golint
	sheet.boundSheet = bs
	sheet.wb = wb
	wb.sheets = append(wb.sheets, sheet)
	return sheet
}

func (wb *Workbook) read(stream []byte) (err error) { // nolint: gocyclo

	var point int32
	var SSTContinue = false
	var sPoint, prevLen int32
	var readType string
	var grbit byte
	var grbitOffset int32

	eof := false

Next:

	recordNumber := stream[point : point+2]
	recordDataLength := int32(helpers.BytesToUint16(stream[point+2 : point+4]))
	sPoint = point + 4

	if bytes.Compare(recordNumber, record.IndexRecord[:]) == 0 {
		_ = new(record.LabelSSt)
		goto EIF
	}

	//BoundSheet

	if bytes.Compare(recordNumber, record.BoundSheetRecord[:]) == 0 {
		var bs record.BoundSheet
		bs.Read(stream[sPoint+grbitOffset : sPoint+recordDataLength])
		//err = binary.Read(bytes.NewBuffer(stream[sPoint:sPoint+recordDataLength]), binary.LittleEndian, &bs)
		_ = wb.addSheet(&bs)
		goto EIF
	}

	//Continue
	if bytes.Compare(recordNumber, record.ContinueRecord[:]) == 0 {

		if SSTContinue {
			readType = "continue"

			prevLen = int32(len(wb.sst.RgbSrc))

			if prevLen != 0 {
				grbitOffset = 1
			} else {
				grbitOffset = 0
			}

			grbit = stream[sPoint]

			wb.sst.RgbSrc = append(wb.sst.RgbSrc, stream[sPoint+grbitOffset:sPoint+recordDataLength]...)
			wb.sst.Read(readType, grbit, prevLen)
		}
		goto EIF
	}

	//SST
	if bytes.Compare(recordNumber, record.SSTRecord[:]) == 0 {
		wb.sst.NewSST(stream[sPoint : sPoint+recordDataLength])

		wb.sst.Read(readType, grbit, prevLen)
		totalSSt := helpers.BytesToUint32(wb.sst.CstTotal[:])
		if recordDataLength >= 8224 || uint32(len(wb.sst.Rgb)) < totalSSt-1 {
			SSTContinue = true
		}
		goto EIF
	}

	if bytes.Compare(recordNumber, record.XFRecord[:]) == 0 {
		xf := new(record.XF)
		xf.Read(stream[sPoint : sPoint+recordDataLength])
		wb.xf=append(wb.xf, *xf)
		goto EIF
	}

	if bytes.Compare(recordNumber, record.FormatRecord[:]) == 0 {
		format := new(record.Format)
		format.Read(stream[sPoint : sPoint+recordDataLength])

		if wb.formats==nil {
			wb.formats = make(map[int]record.Format,0)
		}
		wb.formats[format.GetIndex()]=*format
		goto EIF
	}

	//CodePage
	if bytes.Compare(recordNumber, record.CodePageRecord[:]) == 0 {
		wb.codepage.Read(stream[sPoint : sPoint+recordDataLength])
		goto EIF
	}

	//EOF
	if bytes.Compare(recordNumber, record.EOFRecord[:]) == 0 && recordDataLength == 0 {
		eof = true
	}

	if bytes.Compare(recordNumber, record.BOFMARKS[:]) == 0   {
		copy(wb.vers[:], stream[sPoint : sPoint+2])
		goto EIF
	}

EIF:
	point = point + recordDataLength + 4
	if !eof {
		goto Next
	}

	return err
}
