package record

import "github.com/shakinm/xlsReader/helpers"

//FORMULA: Cell Formula

var FormulaRecord = []byte{0x06, 0x00} // (6h)

/*
A FORMULA record describes a cell that contains a formula.

Record Data
Offset		Field Name		Size		Contents
------------------------------------------------
4			rw				2			Rw
6			col				2			Col
8			ixfe			2			Index to XF record
10			num				8			Current value of the formula
18			grbit			2			Option flags
20			chn				4
24			cce				2			Length of the parsed expression
26			rgce			var			Parsed expression

The chn field should be ignored when you read the BIFF file. If you write a BIFF file,
the chn field must be 00000000h.
The grbit field contains the following option flags:

Bits	Mask	FlagName		Contents
----------------------------------------
0		0001h 	fAlwaysCalc			Always calculate the formula.
1 		0002h 	fCalcOnLoad			Calculate the formula when the file is opened.
2		0004h	(Reserved)
3		0008h	fShrFmla			=1 if the formula is part of shared formula group.
15–4	FFF0h	(Reserved)

The rw field contains the 0-based row number. The col field contains the 0-based
column number.
If the formula evaluates to a number, the num field contains the current calculated
value of the formula in 8-byte IEEE format. If the formula evaluates to a string, a
Boolean value, or an error value, the most significant 2 bytes of the num field are
FFFFh .
A Boolean value is stored in the num field, as shown in the following table. For more
information about Boolean values, see ― "BOOLERR".

Offset		Field Name		Size		Contents
------------------------------------------------
0			otBool			1			=1 always
1			(Reserved)		1			Reserved; must be 0 (zero)
2			f				1			Boolean value
3			(Reserved)		3			Reserved; must be 0 (zero)
6			fExprO			2			=FFFFh

An error value is stored in the num field, as shown in the following table. For more
information about error values, see "BOOLERR"

Offset		Field Name		Size		Contents
------------------------------------------------
0			otErr			1			=2 always
1			(Reserved)		1			Reserved; must be 0 (zero)
2			err				1			Error value
3			(Reserved)		3			Reserved; must be 0 (zero)
6			fExprO			2			=FFFFh

If the formula evaluates to a string, the num field has the structure shown in the
following table.

Offset		Field Name		Size		Contents
------------------------------------------------
0			otString		1			=0 always
1			(Reserved)		5			Reserved; must be 0 (zero)
6			fExprO			2			=FFFFh

The string value is not stored in the num field; instead, it is stored in a STRING
record that immediately follows the FORMULA record.
The cce field contains the length of the formula. The rgce field contains the
formula in its parsed format.

*/

type Formula struct {
	rw    [2]byte
	col   [2]byte
	ixfe  [2]byte
	num   [8]byte
	grbit [2]byte
	chn   [4]byte
	cce   [2]byte
	rgce  []byte
}

func (r *Formula) GetXFIndex() int {
	return int(helpers.BytesToUint16(r.ixfe[:]))
}

func (r *Formula) Read(stream []byte) {
	copy(r.rw[:], stream[:2])
	copy(r.col[:], stream[2:4])
	copy(r.ixfe[:], stream[4:6])
	copy(r.num[:], stream[6:14])
	copy(r.grbit[:], stream[14:16])
	copy(r.chn[:], stream[16:20])
	copy(r.cce[:], stream[20:22])
	copy(r.rgce[:], stream[20:])
}