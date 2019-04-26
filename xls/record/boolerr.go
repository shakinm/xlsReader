package record

import (
	"github.com/shakinm/xlsReader/helpers"
	"reflect"
	"strconv"
)

//BOOLERR: Cell Value, Boolean or Error

var BoolErrRecord = []byte{0x05, 0x02} // (205h)

/*
A BOOLERR record describes a cell that contains a constant Boolean or error value.
The rw field contains the 0-based row number. The col field contains the 0-based
column number.

Record Data
Offset		Field Name		Size		Contents
------------------------------------------------
4			rw				2			Row
6			col				2			Col
8			ixfe			2			Index to the XF record
10			bBoolErr		1			Boolean value or error value
11			fError			1			Boolean/error flag

The bBoolErr field contains the Boolean or error value, as determined by the
fError field. If the fError field contains a 0 (zero), the bBoolErr field contains a
Boolean value; if the fError field contains a 1, the bBoolErr field contains an error
value.
Boolean values are 1 for true and 0 for false.
Error values are listed in the following table.

Error value		Value (hex)		Value (dec.)
--------------------------------------------
#NULL!			00h				0
#DIV/0!			07h				7
#VALUE!			0Fh				15
#REF!			17h				23
#NAME?			1Dh				29
#NUM!			24h				36
#N/A			2Ah				42
*/

type BoolErr struct {
	rw       [2]byte
	col      [2]byte
	ixfe     [2]byte
	bBoolErr [1]byte
	fError   [1]byte
}

func (r *BoolErr) GetRow() [2]byte {
	return r.rw
}

func (r *BoolErr) GetCol() [2]byte {
	return r.col
}

func (r *BoolErr) GetFloat() float64 {

	return float64(r.GetInt64())
}

func (r *BoolErr) GetString() string {
	if int(r.fError[0]) == 1 {
		switch r.GetInt64() {
		case 0:
			return "#NULL!"
		case 7:
			return "#DIV/0!"
		case 15:
			return "#VALUE!"
		case 23:
			return "#REF!"
		case 29:
			return "#NAME?"
		case 36:
			return "#NUM!!"
		case 42:
			return "#N/A"
		}
	} else {
		if r.GetInt64() == 1 {
			return "TRUE"
		} else {
			return "FALSE"
		}
	}
	return strconv.FormatInt(r.GetInt64(), 10)
}

func (r *BoolErr) GetFloat64() (fl float64) {
	return r.GetFloat()
}
func (r *BoolErr) GetInt64() int64 {
	return int64(r.bBoolErr[0])
}

func (r *BoolErr) GetType() string {
	return reflect.TypeOf(r).String()
}

func (r *BoolErr) GetXFIndex() int {
	return int(helpers.BytesToUint16(r.ixfe[:]))
}

func (r *BoolErr) Read(stream []byte) {
	copy(r.rw[:], stream[:2])
	copy(r.col[:], stream[2:4])
	copy(r.ixfe[:], stream[4:6])
	copy(r.bBoolErr[:], stream[6:7])
	copy(r.fError[:], stream[7:8])
}
