package record

// CODEPAGE: Default Code Page
var CodePageRecord = [2]byte{0x42, 0x00} //(42h)

/*
The CODEPAGE record stores the default code page (character set) used when the
workbook was saved.

Record Data
Offset		Field Name		Size		Contents
------------------------------------------------
4 			cv				2 			Code page the file is saved in:
											01B5h (437 dec.) = IBM PC (Multiplan)
											8000h (32768 dec.) = Apple Macintosh
											04E4h (1252 dec.) = ANSI (Microsoft Windows)
*/

type CodePage struct {
	cv [2]byte
}


func (r *CodePage) Read(stream []byte) {
	copy(r.cv[:],stream[:])
}