package record

//BOOLERR: Cell Value, Boolean or Error

var BoolErrRecord = []byte{0x05, 0x02} // (205h)

/*
A BOOLERR record describes a cell that contains a constant Boolean or error value.
The rw field contains the 0-based row number. The col field contains the 0-based
column number.

Record Data
Offset		Field Name		Size		Contents
------------------------------------------------
4			rw				2			Row
6			col				2			Col
8			ixfe			2			Index to the XF record
10			bBoolErr		1			Boolean value or error value
11			fError			1			Boolean/error flag

The bBoolErr field contains the Boolean or error value, as determined by the
fError field. If the fError field contains a 0 (zero), the bBoolErr field contains a
Boolean value; if the fError field contains a 1, the bBoolErr field contains an error
value.
Boolean values are 1 for true and 0 for false.
Error values are listed in the following table.

Error value		Value (hex)		Value (dec.)
--------------------------------------------
#NULL!			00h				0
#DIV/0!			07h				7
#VALUE!			0Fh				15
#REF!			17h				23
#NAME?			1Dh				29
#NUM!			24h				36
#N/A			2Ah				42
*/

type BoolErr struct {
	Rw       [2]byte
	Col      [2]byte
	Ixfe     [2]byte
	BBoolErr [1]byte
	FError   [1]byte
}
