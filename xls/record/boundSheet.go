package record

import (
	"bytes"
	"github.com/shakinm/xlsReader/xls/structure"
	"strings"
)

//This record stores the sheet name, sheet type, and stream position
var BoundSheetRecord = [2]byte{0x85, 0x00} //(85h)

/*
Offset		 Field Name		 Size		 Contents
-------------------------------------------------
4			lbPlyPos		4			Stream position of the start of the BOF record for the sheet
8			grbit			2			Option flags
10			cch				1			Sheet name ( grbit / rgb fields of Unicode String)
11			rgch			var			Sheet name ( grbit / rgb fields of Unicode String)
*/

/*
The grbit structure contains the following options:

Bits	Mask	Option Name		Contents
----------------------------------------
1–0 	0003h 	hsState 		Hidden state:
									00h = visible
									01h = hidden
									02h = very hidden (see text)
7–2 	00FCh 						(Reserved)
15–8	FF00h 	dt				Sheet type:
									00h = worksheet or dialog sheet
									01h = Excel 4.0 macro sheet
									02h = chart
									06h = Visual Basic module
*/

type BoundSheet struct {
	LbPlyPos [4]byte
	Grbit    [2]byte
	Cch      [1]byte
	Rgch     []byte
	stFormat structure.XLUnicodeRichExtendedString
	vers     []byte

}

func (r *BoundSheet) Read(stream []byte, vers []byte) {

	r.vers = vers

	copy(r.LbPlyPos[:], stream[0:4])
	copy(r.Grbit[:], stream[4:6])
	copy(r.Cch[:], stream[6:7])

	if bytes.Compare(vers, FlagBIFF8) == 0 {

		fixedStream:=[]byte{r.Cch[0],0x00}
		fixedStream = append(fixedStream, stream[7:]...)
		_ = r.stFormat.Read(fixedStream)

	} else {
		r.Rgch = make([]byte, int(r.Cch[0]))
		copy(r.Rgch[:], stream[7:])
	}
}
func (r *BoundSheet) GetName() string {
	if bytes.Compare(r.vers, FlagBIFF8) == 0 {
		return r.stFormat.String()
	}
	strLen := int(r.Cch[0])
	return strings.TrimSpace(string(decodeWindows1251(bytes.Trim(r.Rgch[:int(strLen)], "\x00"))))
}