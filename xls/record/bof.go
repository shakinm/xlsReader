package record

//BOF: Beginning of File

var BOFMARKS = []byte{0x09, 0x08} //(809h)

/*
The BOF record marks the beginning of the Book stream in the BIFF file. It also
marks the beginning of record groups (or ―substreams‖ of the Book stream) for
sheets in the workbook. For BIFF2 through BIFF4, the BIFF version is found from the
high-order byte of the record number structure, as shown in the following table. For
BIFF5/BIFF7, and BIFF8 use the vers structure at offset 4 to determine the BIFF version.

Offset		Field Name		Size		Contents
------------------------------------------------
0 			vers 			1			version:
											=00 BIFF2
											=02 BIFF3
											=04 BIFF4
											=08 BIFF5/BIFF7/BIFF8
1			bof				1			09h

*/

type bof struct {
	vers byte
	bof  byte
}
