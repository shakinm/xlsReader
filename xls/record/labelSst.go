package record

import (
	"github.com/shakinm/xlsReader/helpers"
	"reflect"
)

//LABELSST: Cell Value, String Constant/SST

var LabelSStRecord = []byte{0xFD, 0x00} //(FDh)

/*

A LABELSST record describes a cell that contains a string constant from the shared
string table, which is new to BIFF8.

Record Data — BIFF8
Offset		Field Name		Size		Contents
------------------------------------------------
4			rw				2			Row (0-based)
6			col				2			Column (0-based)
8			ixfe			2			Index to the XF record
10			isst			4			Index into the SST record where actual string is stored
*/

type LabelSSt struct {
	rw   [2]byte
	col  [2]byte
	ixfe [2]byte
	isst [4]byte
	sst  *SST
}

func (r *LabelSSt) GetRow() [2]byte {
	return r.rw
}

func (r *LabelSSt) GetCol() [2]byte {
	return r.col
}

func (r *LabelSSt) GetString() string {
	index := helpers.BytesToUint32(r.isst[:])
	if index < uint32(len(r.sst.Rgb)) {
		return r.sst.Rgb[index].String()
	}
	return ""
}

func (r *LabelSSt) GetFloat64() (fl float64) {
	return fl
}
func (r *LabelSSt) GetInt64() (in int64) {
	return in
}

func (r *LabelSSt) GetType() string {
	return reflect.TypeOf(r).String()
}

func (r *LabelSSt) GetXFIndex() int {
	return int(helpers.BytesToUint16(r.ixfe[:]))
}

func (r *LabelSSt) Read(stream []byte, sst *SST) {
	r.sst = sst
	copy(r.rw[:], stream[:2])
	copy(r.col[:], stream[2:4])
	copy(r.ixfe[:], stream[4:6])
	copy(r.isst[:], stream[6:10])
}
