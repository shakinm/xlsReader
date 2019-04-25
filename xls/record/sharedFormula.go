package record

// SHRFMLA: Shared Formula

var SharedFormulaRecord = []byte{0xBC, 0x00} // (BCh)

/*
The SHRFMLA record is a file size optimization. It is used with the FORMULA record to
compress the amount of storage required for the parsed expression ( rgce ). In
earlier versions of Excel, if you read a FORMULA record in which the rgce field
contained a ptgExp parse token, the FORMULA record contained an array formula.
In Excel 5.0 and later, this could indicate either an array formula or a shared
formula.
If the record following the FORMULA is an ARRAY record, the FORMULA record
contains an array formula. If the record following the FORMULA is a SHRFMLA record,
the FORMULA record contains a shared formula. You can also test the fShrFmla bit
in the FORMULA record‘s grbit field to determine this.
When reading a file, you must convert the FORMULA and SHRFMLA records to an
equivalent FORMULA record if you plan to use the parsed expression. To do this,
take all of the FORMULA record up to (but not including) the cce field, and then
append to that the SHRFMLA record from its cce field to the end. You must then
convert some ptg s; this is explained later in this article.
Following the SHRFMLA record are one or more FORMULA records containing ptgExp
tokens that have the same rwFirst and colFirst fields as those in the ptgExp in
the first FORMULA . There is only one SHRFMLA record for each shared-formula
record group.
To convert the ptg s, search the rgce field from the SHRFMLA record for any
ptgRefN , ptgRefNV , ptgRefNA , ptgAreaN , ptgAreaNV , or ptgAreaNA tokens.
Add the corresponding FORMULA record‘s rw and col fields to the rwFirst and
colFirst fields in the ptg s from the SHRFMLA . Finally, convert the ptg s as shown
in the following table.

Convert
this ptg			To this ptg
-------------------------------
ptgRefN				ptgRef
ptgRefNV			ptgRefV
ptgRefNA			ptgRefA
ptgAreaNV			ptgArea
ptgAreaNA			ptgAreaA

Remember that STRING records can appear after FORMULA records if the formula
evaluates to a string.
If your code writes a BIFF file, always write standard FORMULA records; do not
attempt to use the SHRFMLA optimization.

Record Data
Offset		Name		Size		Contents
--------------------------------------------
4			rwFirst		2			First row
6			rwLast		2			Last row
8			colFirst	1			First column
9			colLast		1			Last column
10			(Reserved)	2
12			cce			2			Length of the parsed expression
14			rgce		var			Parsed expression

*/

type ShareFormula struct {
	RwFirst  [2]byte
	RwLast   [2]byte
	ColFirst [1]byte
	ColLast  [1]byte
	Reserved [2]byte
	Cce      [2]byte
	Rgce     []byte
}
