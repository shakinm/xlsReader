package record

// ROW: Describes a Row

var RowRecord = []byte{0x08, 0x02} // (208h)

/*
A ROW record describes a single row on an Excel sheet. ROW records and their
associated cell records occur in blocks of up to 32 rows. Each block ends with a
DBCELL record.

Record Data
Offset		Name		Size		Contents
--------------------------------------------
4 			rw			2			Row number.
6			colMic		2			First defined column in the row.
8			colMac		2			Last defined column in the row, plus 1.
10			miyRw		2			Row height.
12			irwMac		2			Used by Excel to optimize loading the file; if you are creating a BIFF file, set irwMac to 0.
14 			(Reserved) 	2
16			grbit		2			Option flags.
18			ixfe		2			If fGhostDirty=1 (see grbit structure), this is the index to the XF record for the row.
									Otherwise, this structure is undefined.
									Note: ixfe uses only the low-order 12 bits of the structure
									(bits 11–0). Bit 12 is fExAsc , bit 13 is fExDsc , and bits
									14 and 15 are reserved. fExAsc and fExDsc are set to
									true if the row has a thick border on top or on bottom,
									respectively.

The grbit structure contains the following option flags:
Offset		Bits		Mask		Name			Contents
--------------------------------------------------------
0			2–0			07h			iOutLevel		Outline level of the row
			3			08h			(Reserved)
			4			10h			fCollapsed		=1 if the row is collapsed in outlining
			5			20h			fDyZero			=1 if the row height is set to 0 (zero)
			6			40h			fUnsynced		=1 if the font height and row height are not compatible
			7			80h			fGhostDirty		=1 if the row has been formatted, even if it contains all blank cells
1			7–0			FFh			(Reserved)

The rw structure contains the 0-based row number. The colMic and colMac fields give
the range of defined columns in the row.

The miyRw structure contains the row height, in units of 1/20 th of a point. The miyRw
structure may have the 8000h (2 15 ) bit set, indicating that the row is standard height.
The low-order 15 bits must still contain the row height. If you hide the row — either
by setting row height to 0 (zero) or by using the Hide command — miyRw still
contains the original row height. This allows Excel to restore the original row height
when you click the Unhide button.
Each row can have default cell attributes that control the format of all undefined cells
in the row. By specifying default cell attributes for a particular row, you are
effectively formatting all the undefined cells in the row without using memory for
those cells. Default cell attributes do not affect the formats of cells that are explicitly
defined.
For example, if you want all of row 3 to be left-aligned, you could define all 256 cells
in the row and specify that each individual cell be left-aligned. This would require
storage for each of the 256 cells. An easy alternative would be to set the default cell
for row 3 to be left-aligned and not define any individual cells in row 3.

*/

type Row struct {
	Rw       [2]byte
	ColMic   [2]byte
	ColMac   [2]byte
	MiyRw    [2]byte
	IrwMac   [2]byte
	Reserved [2]byte
	Grbit    [2]byte
	Ixfe     [2]byte
}
