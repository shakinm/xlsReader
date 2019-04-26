package record

import (
	"encoding/binary"
	"github.com/shakinm/xlsReader/helpers"
	"math"
	"reflect"

	"strconv"
)

//NUMBER: Cell Value, Floating-Point Number

var NumberRecord = []byte{0x03, 0x02} //(203h)
/*
A NUMBER record describes a cell containing a constant floating-point number. The
rw field contains the 0-based row number. The col field contains the 0-based
column number. The number is contained in the num field in 8-byte IEEE floating-
point format.

Record Data
Offset		Name		Size		Contents
--------------------------------------------
4			rw			2			Row
6			col			2			Column
8			ixfe		2			Index to the XF record
10			num			8			Floating-point number value
*/

type Number struct {
	rw   [2]byte
	col  [2]byte
	ixfe [2]byte
	num  [8]byte
}

func (r *Number) GetRow() [2]byte {
	return r.rw
}

func (r *Number) GetCol() [2]byte {
	return r.col
}

func (r *Number) GetFloat() float64 {
	bits := binary.LittleEndian.Uint64(r.num[:])
	float := math.Float64frombits(bits)
	return float
}

func (r *Number) GetString() string {

	return strconv.FormatFloat(r.GetFloat(), 'f', 6, 64)
}

func (r *Number) GetFloat64() (fl float64) {
	return r.GetFloat()
}
func (r *Number) GetInt64() (in int64) {
	return in
}

func (r *Number) GetType() string {
	return reflect.TypeOf(r).String()
}

func (r *Number) GetXFIndex() int {
	return int(helpers.BytesToUint16(r.ixfe[:]))
}

func (r *Number) Read(stream []byte) {
	copy(r.rw[:], stream[:2])
	copy(r.col[:], stream[2:4])
	copy(r.ixfe[:], stream[4:6])
	copy(r.num[:], stream[6:14])
}
