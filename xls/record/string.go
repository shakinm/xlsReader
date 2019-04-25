package record

// STRING: String Value of a Formula

var StringRecord = []byte{0x07, 0x02} // (207h)

/*
When a formula evaluates to a string, a STRING record occurs after the FORMULA
record. If the formula is part of an array, the STRING record occurs after the ARRAY
record.

Record Data
Offset		Name		Size		Contents
--------------------------------------------
4			cch			2			Length of the string
6			grbit		1			0= Compressed unicode string
									1= Uncompressed unicode string
7			rgch		var			String
*/

type String struct {
	Cch   [2]byte
	Grbit [2]byte
	Rgch  []byte

}