package record

import (
	"github.com/shakinm/xlsReader/helpers"
	"github.com/shakinm/xlsReader/xls/structure"
	"reflect"
)

// RK: Cell Value, RK Number

var RkRecord = []byte{0x7E, 0x02} //(7Eh)
/*
Excel uses an internal number type, called an RK number, to save memory and disk
space.

Record Data
Offset		Name		Size		Contents
--------------------------------------------
4			rw			2			Row
6			col			2			Column
8			ixfe		2			Index to the XF record
10			rk			4			RK number (see the following description)

An RK number is either a 30-bit integer or the most significant 30 bits of an IEEE
number. The two LSBs of the 32-bit rk field are always reserved for RK type
encoding; this is why the RK numbers are 30 bits, not the full 32.

*/

type Rk struct {
	rw   [2]byte
	col  [2]byte
	ixfe [2]byte
	rk   structure.RKNum
}

func (r *Rk) GetRow() [2]byte {
	return r.rw
}

func (r *Rk) GetCol() [2]byte {
	return r.col
}

func (r *Rk) GetFloat64() float64 {
	return r.rk.GetFloat()
}

func (r *Rk) GetInt64() int64 {
	return r.rk.GetInt64()
}

func (r *Rk) GetType() string {
	return reflect.TypeOf(r).String()
}

func (r *Rk) GetString() (s string) {
	return r.rk.GetString()
}

func (r *Rk) GetXFIndex() int {
	return int(helpers.BytesToUint16(r.ixfe[:]))
}

func (r *Rk) Read(stream []byte) {
	copy(r.rw[:], stream[:2])
	copy(r.col[:], stream[2:4])
	copy(r.ixfe[:], stream[4:6])
	copy(r.rk[:], stream[6:10])
}

func (r *Rk) Get() *Rk {
	return r
}
