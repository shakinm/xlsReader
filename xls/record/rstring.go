package record

// RSTRING: Cell with Character Formatting

var RStringRecord = []byte{0xD6, 0x00} // (D6h)

/*
When part of a string in a cell has character formatting, an RSTRING record is
written instead of the LABEL record. The RSTRING record is obsolete in BIFF8,
replaced by the LABELSST and SST records.

Record Data
Offset		Name		Size		Contents
--------------------------------------------
4			rw			2			Row
6			col			2			Column
8			ixfe		2			Index to the XF record
10			cch			2			Length of the string
12			rgch		var			String
var			cruns		1			Count of STRUN structures
var			rgstrun		var			Array of STRUN structures

The STRUN structure contains formatting information about the string. A STRUN
structure occurs every time the text formatting changes. The STRUN structure is
described in the following table.

Offset		Name		Size		Contents
--------------------------------------------
0			ich			1			Index to the first character to which the formatting applies
1			ifnt		1			Index to the FONT record

*/

type RString struct {
	Rw      [2]byte
	Col     [2]byte
	Ixfe    [2]byte
	Cch     [2]byte
	Rgch    []byte
	Cruns   [1]byte
	Rgstrun []byte
}
