package record

import (
	"github.com/Alliera/xlsReader/helpers"
	"golang.org/x/text/encoding/charmap"
	"reflect"
	"strings"
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

type LabelBIFF8 struct {
	rw    [2]byte
	col   [2]byte
	ixfe  [2]byte
	cch   [2]byte
	grbit [1]byte
	rgb   []byte
}

type LabelBIFF5 struct {
	rw    [2]byte
	col   [2]byte
	ixfe  [2]byte
	cch   [2]byte
	grbit [1]byte
	rgb   []byte
}

func (r *LabelBIFF8) GetRow() [2]byte {
	return r.rw
}

func (r *LabelBIFF8) GetCol() [2]byte {
	return r.col
}

func (r *LabelBIFF8) GetString() string {
	if int(r.grbit[0]) == 1 {
		name := helpers.BytesToUints16(r.rgb[:])
		runes := utf16.Decode(name)
		return string(runes)
	} else {
		return string(decodeWindows1251(r.rgb[:]))
	}
}

func (r *LabelBIFF8) GetFloat64() (fl float64) {
	return fl
}
func (r *LabelBIFF8) GetInt64() (in int64) {
	return in
}

func (r *LabelBIFF8) GetType() string {
	return reflect.TypeOf(r).String()
}

func (r *LabelBIFF8) GetXFIndex() int {
	return int(helpers.BytesToUint16(r.ixfe[:]))
}

func (r *LabelBIFF8) Read(stream []byte) {

	copy(r.rw[:], stream[:2])
	copy(r.col[:], stream[2:4])
	copy(r.ixfe[:], stream[4:6])
	copy(r.cch[:], stream[6:8])
	copy(r.grbit[:], stream[8:9])
	if int(r.grbit[0]) == 1 {
		r.rgb = make([]byte, helpers.BytesToUint16(r.cch[:])*2)
	} else {
		r.rgb = make([]byte, helpers.BytesToUint16(r.cch[:]))
	}

	copy(r.rgb[:], stream[9:])
}

func (r *LabelBIFF5) GetRow() [2]byte {
	return r.rw
}

func (r *LabelBIFF5) GetCol() [2]byte {
	return r.col
}

func (r *LabelBIFF5) GetString() string {
	strLen := helpers.BytesToUint16(r.cch[:])
	return strings.TrimSpace(string(decodeWindows1251(r.rgb[:int(strLen)])))
}

func (r *LabelBIFF5) GetFloat64() (fl float64) {
	return fl
}
func (r *LabelBIFF5) GetInt64() (in int64) {
	return in
}

func (r *LabelBIFF5) GetType() string {
	return reflect.TypeOf(r).String()
}

func (r *LabelBIFF5) GetXFIndex() int {
	return int(helpers.BytesToUint16(r.ixfe[:]))
}

func (r *LabelBIFF5) Read(stream []byte) {

	copy(r.rw[:], stream[:2])
	copy(r.col[:], stream[2:4])
	copy(r.ixfe[:], stream[4:6])
	copy(r.cch[:], stream[6:8])
	//copy(r.grbit[:], stream[8:9])
	r.rgb = make([]byte, helpers.BytesToUint16(r.cch[:]))
	copy(r.rgb[:], stream[8:])
}

func decodeWindows1251(ba []uint8) []uint8 {
	dec := charmap.Windows1251.NewDecoder()
	out, _ := dec.Bytes(ba)
	return out
}
