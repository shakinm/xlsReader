package record

import (
	"github.com/shakinm/xlsReader/helpers"
	"reflect"
	"unicode/utf16"
)

//LABEL: Cell Value, String Constant (204h)

var LabelRecord = []byte{0x04, 0x02} //(204h)

/*

A LABEL record describes a cell that contains a pre-BIFF8 string constant. Note:
this was replaced in BIFF8 by LABELSST .

Record Data 8
Offset		Field Name		Size		Contents
------------------------------------------------
4			rw				2			Row (0-based)
6			col				2			Column (0-based)
8			ixfe			2			Index to the XF record
10			cch				2			Length of the string (must be <= 255)
12			grbit			1			Option flags
13			rgb				var			Array of string characters
*/

type Label struct {
	rw    [2]byte
	col   [2]byte
	ixfe  [2]byte
	cch   [2]byte
	grbit [1]byte
	rgb   []byte
}

func (r *Label) GetRow() [2]byte {
	return r.rw
}

func (r *Label) GetCol() [2]byte {
	return r.col
}

func (r *Label) GetString() string {
	//if 	r.grbit[:]<<=1 {
		name := helpers.BytesToUints16(r.rgb[:])
		runes := utf16.Decode(name)
		return string(runes)
	//} else {
	//
	//	return string(r.rgb[:])
	//}
}

func (r *Label) GetFloat64() (fl float64) {
	return fl
}
func (r *Label) GetInt64() (in int64) {
	return in
}

func (r *Label) GetType() string {
	return reflect.TypeOf(r).String()
}

func (r *Label) GetXFIndex() int {
	return int(helpers.BytesToUint16(r.ixfe[:]))
}

func (r *Label) Read(stream []byte) {

	copy(r.rw[:], stream[:2])
	copy(r.col[:], stream[2:4])
	copy(r.ixfe[:], stream[4:6])
	copy(r.cch[:], stream[6:8])
	copy(r.grbit[:], stream[8:9])
	if int(r.grbit[0]) ==1 {
		r.rgb=make([]byte, helpers.BytesToUint16(r.cch[:])*2)
	} else {
		r.rgb=make([]byte, helpers.BytesToUint16(r.cch[:]))
	}

	copy(r.rgb[:], stream[9:])
}
