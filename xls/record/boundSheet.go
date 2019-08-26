package record

import (
	"github.com/shakinm/xlsReader/helpers"
	"github.com/shakinm/xlsReader/xls/structure"
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
	Rgch     structure.XLUnicodeRichExtendedString
}

func (r *BoundSheet) Read(stream []byte) {

	var rgbSize uint8
	oft := uint32(7)

	copy(r.LbPlyPos[:], stream[0:4])
	copy(r.Grbit[:], stream[4:6])
	copy(r.Cch[:], stream[6:7])

	//offset 2
	r.Rgch.FHighByte = stream[iOft(&oft, 0):iOft(&oft, 1)][0]

	if r.Rgch.FHighByte&1 == 1 {
		rgbSize = uint8(r.Cch[0]) * 2

	} else {
		rgbSize = uint8(r.Cch[0])

	}

	if r.Rgch.FHighByte>>3&1 == 1 { // if fRichSt == 1

		copy(r.Rgch.CRun[:], stream[iOft(&oft, 0):iOft(&oft, 2)])
	}

	if r.Rgch.FHighByte>>2&1 == 1 { //fExtSt  == 1
		//offset 4
		copy(r.Rgch.CbExtRst[:], stream[iOft(&oft, 0):iOft(&oft, 4)])
	}

	//offset rgbSize
	r.Rgch.Rgb = make([]byte, uint32(rgbSize))
	copy(r.Rgch.Rgb[0:], stream[iOft(&oft, 0):iOft(&oft, uint32(rgbSize))])

	if r.Rgch.FHighByte>>3&1 == 1 { // if fRichSt == 1
		cRunSize := helpers.BytesToUint16(r.Rgch.CRun[:])
		for i := uint16(0); i <= cRunSize-1; i++ {
			var rgRun structure.FormatRun
			copy(rgRun.Ich[:], stream[iOft(&oft, 0):iOft(&oft, 2)])
			copy(rgRun.Ifnt.Ifnt[:], stream[iOft(&oft, 0):iOft(&oft, 2)])
			r.Rgch.RgRun = append(r.Rgch.RgRun, rgRun)
		}
	}

	if r.Rgch.FHighByte>>2&1 == 1 { //fExtSt  == 1
		copy(r.Rgch.ExtRst.Reserved[:], stream[iOft(&oft, 0):iOft(&oft, 2)])
		copy(r.Rgch.ExtRst.Cb[:], stream[iOft(&oft, 0):iOft(&oft, 2)])

		copy(r.Rgch.ExtRst.Phs.Ifnt.Ifnt[:], stream[iOft(&oft, 0):iOft(&oft, 2)])
		copy(r.Rgch.ExtRst.Phs.Info[:], stream[iOft(&oft, 0):iOft(&oft, 2)])

		copy(r.Rgch.ExtRst.Rphssub.Crun[:], stream[iOft(&oft, 0):iOft(&oft, 2)])
		copy(r.Rgch.ExtRst.Rphssub.Cch[:], stream[iOft(&oft, 0):iOft(&oft, 2)])

		copy(r.Rgch.ExtRst.Rphssub.St.CchCharacters[:], stream[iOft(&oft, 0):iOft(&oft, 2)])

		rgchDataSize := helpers.BytesToUint16(r.Rgch.ExtRst.Rphssub.St.CchCharacters[:]) * 2
		for i := uint16(0); i <= rgchDataSize; i++ {
			r.Rgch.ExtRst.Rphssub.St.RgchData = append(r.Rgch.ExtRst.Rphssub.St.RgchData, stream[iOft(&oft, 0):iOft(&oft, 2)]...)
		}

		//The number of elements in this array is rphssub.crun
		phRunsSizeL := helpers.BytesToUint16(r.Rgch.ExtRst.Rphssub.Crun[:])
		if phRunsSizeL > 0 {
			for i := uint16(0); i <= phRunsSizeL; i++ {
				var phRuns structure.PhRuns
				copy(phRuns.IchFirst[:], stream[iOft(&oft, 0):iOft(&oft, 2)])
				copy(phRuns.IchMom[:], stream[iOft(&oft, 0):iOft(&oft, 2)])
				copy(phRuns.CchMom[:], stream[iOft(&oft, 0):iOft(&oft, 2)])

				r.Rgch.ExtRst.Rgphruns = append(r.Rgch.ExtRst.Rgphruns, phRuns)
			}
		}
	}

}
