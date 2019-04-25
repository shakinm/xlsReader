package record

import (
	"encoding/binary"
	"github.com/shakinm/xlsReader/helpers"
	"github.com/shakinm/xlsReader/xls/structure"
)

// MULRK: Multiple RK Cells

var MulRKRecord = []byte{0xBD, 0x00} // (BDh)

/*
The MULRK record stores up to the equivalent of 256 RK records; the MULRK record is
a file size optimization. The number of 6-byte RKREC structures can be determined
from the ColLast field and is equal to (colLast-colFirst+1) . The maximum
length of the MULRK record is (256x6+10)=1546 bytes, because Excel has at most
256 columns. Note: storing 256 RK numbers in the MULRK record takes 1,546 bytes
as compared with 3,584 bytes for 256 RK records.

Record Data
Offset		Name		Size		Contents
--------------------------------------------
4			rw			2			Row number (0-based)
6			colFirst	2			Column number (0-based) of the first column of the
									multiple RK record
8			rgrkrec		var			Array of 6-byte RKREC structures
10			colLast		2			Last column containing the RKREC structure
*/

type MulRk struct {
	rw       [2]byte
	colFirst [2]byte
	rgrkrec  []structure.RKREC
	colLast  [2]byte
}

func (r *MulRk) GetArrayRKRecord() (rkRecords []Rk) {

	for k, rkrec := range r.rgrkrec {
		var rk Rk
		rk.rw = r.rw
		binary.LittleEndian.PutUint16(rk.col[:], uint16(k)+helpers.BytesToUint16(r.colFirst[:]))
		rk.ixfe = rkrec.Ixfe
		rk.rk = rkrec.RK
		rkRecords = append(rkRecords, rk)
	}

	return
}

func (r *MulRk) Read(stream []byte) {
	copy(r.rw[:], stream[:2])
	copy(r.colFirst[:], stream[2:4])
	copy(r.colLast[:], stream[len(stream)-2:])

	cf := helpers.BytesToUint16(r.colFirst[:])
	cl := helpers.BytesToUint16(r.colLast[:])
	for i := 0; i <= int(cl-cf); i++ {
		sPoint := 4 + (i * 6)
		var rkRec structure.RKREC
		copy(rkRec.Ixfe[:], stream[sPoint:sPoint+2])
		copy(rkRec.RK[:], stream[sPoint+2:sPoint+6])
		r.rgrkrec = append(r.rgrkrec, rkRec)
	}

}
