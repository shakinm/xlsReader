package record

// CONTINUE: Continues Long Records
var ContinueRecord = [2]byte{0x3c, 0x00} //(0x3c)

/*
Records longer than 8,228 bytes (2,084 bytes in BIFF7 and earlier) must be split into
several records. The first section appears in the base record; subsequent sections
appear in CONTINUE records.
In BIFF8, the TXO record is always followed by CONTINUE records that store the
string data and formatting runs.

Record Data
Offset		Name	Size	Contents
------------------------------------
4 					var 	Continuation of record data

If the continued data is a string, the CONTINUE record also has a structure to indicate
whether the string is compressed or uncompressed unicode.

Record Data
Offset		Field Name		Size	Contents
--------------------------------------------
4 			grbit			1 		0= Compressed unicode string
									1= Uncompressed unicode string
5							var		Continuation of record data

*/

type Continue struct {
	data []byte
}