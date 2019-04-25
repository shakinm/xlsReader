package record

import "github.com/shakinm/xlsReader/helpers"

//INDEX: Index Record

var IndexRecord = [2]byte{0x02, 0x0B} // (20Bh)

/*
Excel writes an INDEX record immediately after the BOF record for each worksheet
substream in a BIFF file. For more information about the INDEX record

Record Data — BIFF8
Offset		Field Name		Size		Contents
------------------------------------------------
4			(Reserved)		4			Reserved; must be 0 (zero)
8			rwMic			4			First row that exists on the sheet
12			rwMac			4			Last row that exists on the sheet, plus 1
16			(Reserved)		4			Reserved; must be 0 (zero)
20			rgibRw			var			Array of file offsets to the DBCELL records for each
										block of ROW records. A block contains ROW records for up to 32 rows.

Record Data — BIFF7
Offset		Field Name		Size		Contents
------------------------------------------------
4			(Reserved)		4			Reserved; must be 0 (zero)
8			rwMic			2			First row that exists on the sheet
10			rwMac			2			Last row that exists on the sheet, plus 1
12			(Reserved)		4			Reserved; must be 0 (zero)
16			rgibRw			var			Array of file offsets to the DBCELL records for each
										block of ROW records. A block contains ROW records for up to 32 rows.

The rwMic field contains the number of the first row in the sheet that contains a
value or a formula that is referenced by a cell in some other row. Because rows (and
columns) are always stored 0-based rather than 1-based (as they appear on the
screen), cell A1 is stored as row 0, cell A2 is row 1, and so on. The rwMac field
contains the 0-based number of the last row in the sheet, plus 1.

*/

type Index struct {
	reserved  [4]byte
	rwMic     [4]byte
	rwMac     [4]byte
	reserved2 [4]byte
	rgibRw    []byte
}

func (r *Index) GetMaxRow() uint32 {
	return helpers.BytesToUint32(r.rwMac[0:0]) + 1
}

func (r *Index) Read(stream []byte) {
	copy(r.reserved[:], stream[:4])
	copy(r.rwMic[:], stream[4:8])
	copy(r.rwMac[:], stream[8:12])
	copy(r.reserved2[:], stream[12:16])
	copy(r.rgibRw, stream[16:])
}
