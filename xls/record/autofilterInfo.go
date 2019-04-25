package record

import "github.com/shakinm/xlsReader/helpers"

//AUTOFILTERINFO: Drop-Down Arrow Count

var AutofilterInfoRecord = [2]byte{0x9D, 0x00} //(9Dh)

/*
This record stores the count of AutoFilter drop-down arrows. Each drop-down arrow
has a corresponding OBJ record. If at least one AutoFilter is active (in other words,
the range was filtered at least once), there is a corresponding FILTERMODE record in
the file. There is also one AUTOFILTER record for each active filter.

Record Data
Offset		Field Name		Size		Contents
------------------------------------------------
4			cEntries		2			Number of AutoFilter drop-down arrows on the sheet

*/

type AutofilterInfo struct {
	cEntries [2]byte
}

func (r *AutofilterInfo) GetCountEntries() uint16 {
	return helpers.BytesToUint16(r.cEntries[:])
}

func (r *AutofilterInfo) Read(stream []byte) {
	copy(r.cEntries[:], stream[:2])
}
