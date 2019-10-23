package record

import (
	"encoding/binary"
	"github.com/shakinm/xlsReader/helpers"
)

//MULBLANK: Multiple Blank Cells

var MulBlankRecord = []byte{0xBE, 0x00} // (BEh)

/*
The MULBLANK record stores up to the equivalent of 256 BLANK records; the
MULBLANK record is a file size optimization. The number of ixfe fields can be
etermined from the ColLast field and is equal to ( colLast-colFirst+1 ). The
maximum length of the MULBLANK record is (256x2+10)=522 bytes, because Excel
can have at most 256 columns. Note: storing 256 blank cells in the MULBLANK
record takes 522 bytes as compared with 2,560 bytes for 256 BLANK records.

Record Data
Offset		Name		Size		Contents
--------------------------------------------
4			rw			2			Row number (0-based)
6			colFirst	2			Column number (0-based) of the first column of the
									multiple RK record
8			rgixfe		var			Array of indexes to XF records
10			colLast		2			Last column containing the BLANKREC structure
*/

type MulBlank struct {
	rw       [2]byte
	colFirst [2]byte
	rgixfe   [][2]byte
	colLast  [2]byte
}


func (r *MulBlank) GetArrayBlRecord() (blkRecords []Blank) {

	for k, rgixfe := range r.rgixfe {
		var bl Blank
		bl.rw = r.rw
		binary.LittleEndian.PutUint16(bl.col[:], uint16(k)+helpers.BytesToUint16(r.colFirst[:]))
		bl.ixfe=rgixfe
		blkRecords= append(blkRecords, bl)
	}

	return
}

func (r *MulBlank) Read(stream []byte) {
	copy(r.rw[:], stream[:2])
	copy(r.colFirst[:], stream[2:4])
	copy(r.colLast[:], stream[len(stream)-2:])

	cf := helpers.BytesToUint16(r.colFirst[:])
	cl := helpers.BytesToUint16(r.colLast[:])
	for i := 0; i <= int(cl-cf); i++ {
		sPoint := 4 + (i * 6)
		var indexXF [2]byte
		copy(indexXF[:], stream[sPoint:sPoint+2])
		r.rgixfe = append(r.rgixfe, indexXF)
	}

}
