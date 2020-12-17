package record

import (
	"github.com/Alliera/xlsReader/helpers"
	"reflect"
)

//BLANK: Cell Value, Blank Cell

var BlankRecord = []byte{0x01, 0x02} //(201h)

/*
A BLANK record describes an empty cell. The rw field contains the 0-based row
number. The col field contains the 0-based column number.

Record Data
Offset		Name		Size		Contents
--------------------------------------------
4			rw			2			Row
6			col			2			Column
8			ixfe		2			Index to the XF record
*/

type Blank struct {
	rw   [2]byte
	col  [2]byte
	ixfe [2]byte
}

func (r *Blank) GetRow() [2]byte {
	return r.rw
}

func (r *Blank) GetCol() [2]byte {
	return r.col
}

func (r *Blank) Read(stream []byte) {
	copy(r.rw[:], stream[:2])
	copy(r.col[:], stream[2:4])
	copy(r.ixfe[:], stream[4:6])
}

func (r *Blank) GetString() (str string) {
	return str
}

func (r *Blank) GetFloat64() (fl float64) {
	return fl
}
func (r *Blank) GetInt64() (in int64) {
	return in
}

func (r *Blank) GetType() string {
	return reflect.TypeOf(r).String()
}

func (r *Blank) GetXFIndex() int {
	return int(helpers.BytesToUint16(r.ixfe[:]))
}

func (r *Blank) Get() *Blank {
	return r
}
