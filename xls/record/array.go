package record

//ARRAY: Array-Entered Formula

var ArrayRecord = []byte{0x021, 0x02} //(221h)

/*
An ARRAY record describes a formula that was array-entered into a range of cells.
The range of cells in which the array is entered is defined by the rwFirst , rwLast ,
colFirst , and colLast fields.
The ARRAY record occurs directly after the FORMULA record for the cell in the upper-
left corner of the array — that is, the cell defined by the rwFirst and colFirst
fields.

Record Data
Offset		Field Name		Size		Contents
------------------------------------------------
4			rwFirst			2			First row of the array
6			rwLast			2			Last row of the array
8			colFirst		1			First column of the array
9			colLast			1			Last column of the array
10			grbit			2			Option flags
12			chn				4
16			cce				2			Length of the parsed expression
18			rgce			var			Parsed formula expression

Ignore the chn field when reading the BIFF file. If a BIFF file is written, the chn field
must be 00000000h .

The grbit field contains the following option flags:

Offset		Bits		Mask		Flag Name		Contents
------------------------------------------------------------
0			0			01h			fAlwaysCalc		Always calculate the formula.
 			1			02h			fCalcOnLoad		Calculate the formula when the file is opened.
			7–2			FCh			(unused)
1			7–2			FFh			(unused)

*/

type Array struct {
	RwFirst  [2]byte
	RwLast   [2]byte
	ColFirst [1]byte
	ColLast  [1]byte
	Grbit    [2]byte
	Chn      [4]byte
	Cce      [2]byte
	Rgce     []byte
}
