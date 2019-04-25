package structure

import (
	"unicode/utf16"
	"github.com/shakinm/xlsReader/helpers"
)

type XLUnicodeRichExtendedString struct {
	Cch [2]byte

	/*
		A - fHighByte (1 bit): A bit that specifies whether the characters in rgb are double-byte characters.
		MUST be a value from the following table:

		B - reserved1 (1 bit): MUST be zero, and MUST be ignored.
		C - fExtSt (1 bit): A bit that specifies whether the string contains phonetic string data
		D - fRichSt (1 bit): A bit that specifies whether the string is a rich string and the string
			has at least
		reserved2 (4 bits): MUST be zero, and MUST be ignored.
	*/
	FHighByte byte // ABCD
	CRun      [2]byte
	CbExtRst  [4]byte
	Rgb       []byte // If fHighByte is 0x0 size = cch.  If fHighByte is 0x1 size = cch*2

	/*
		An optional array of FormatRun structure that specifies the formatting for each
		text run. The number of elements in the array is cRun. MUST exist if and only if fRichSt is 0x1.
	*/
	RgRun []FormatRun

	/*
		An optional ExtRst that specifies the phonetic string data. The size of this structure is
		cbExtRst. MUST exist if and only if fExtSt is 0x1.
	*/
	ExtRst ExtRst
}

func (s *XLUnicodeRichExtendedString) Read(stream []byte) {
	var rgbSize uint16
	oft := uint32(0)

	copy(s.Cch[:], stream[iOft(&oft, 0):iOft(&oft, 2)])

	//offset 2
	s.FHighByte = stream[iOft(&oft, 0):iOft(&oft, 1)][0]

	if s.FHighByte&1 == 1 {
		rgbSize = helpers.BytesToUint16(s.Cch[:])*2

	} else {
		rgbSize = helpers.BytesToUint16(s.Cch[:])

	}

	if s.FHighByte>>3&1 == 1 { // if fRichSt == 1

		copy(s.CRun[:], stream[iOft(&oft, 0):iOft(&oft, 2)])
	}

	if s.FHighByte>>2&1 == 1 { //fExtSt  == 1
		//offset 4
		copy(s.CbExtRst[:], stream[iOft(&oft, 0):iOft(&oft, 4)])
	}

	//offset rgbSize
	s.Rgb = make([]byte, uint32(rgbSize))
	copy(s.Rgb[0:], stream[iOft(&oft, 0):iOft(&oft, uint32(rgbSize))])

	if s.FHighByte>>3&1 == 1 { // if fRichSt == 1
		cRunSize := helpers.BytesToUint16(s.CRun[:])
		for i := uint16(0); i <= cRunSize-1; i++ {
			var rgRun FormatRun
			copy(rgRun.Ich[:], stream[iOft(&oft, 0):iOft(&oft, 2)])
			copy(rgRun.Ifnt.Ifnt[:], stream[iOft(&oft, 0):iOft(&oft, 2)])
			s.RgRun = append(s.RgRun, rgRun)
		}
	}

	if s.FHighByte>>2&1 == 1 { //fExtSt  == 1
		copy(s.ExtRst.Reserved[:], stream[iOft(&oft, 0):iOft(&oft, 2)])
		copy(s.ExtRst.Cb[:], stream[iOft(&oft, 0):iOft(&oft, 2)])

		copy(s.ExtRst.Phs.Ifnt.Ifnt[:], stream[iOft(&oft, 0):iOft(&oft, 2)])
		copy(s.ExtRst.Phs.Info[:], stream[iOft(&oft, 0):iOft(&oft, 2)])

		copy(s.ExtRst.Rphssub.Crun[:], stream[iOft(&oft, 0):iOft(&oft, 2)])
		copy(s.ExtRst.Rphssub.Cch[:], stream[iOft(&oft, 0):iOft(&oft, 2)])

		copy(s.ExtRst.Rphssub.St.CchCharacters[:], stream[iOft(&oft, 0):iOft(&oft, 2)])

		rgchDataSize := helpers.BytesToUint16(s.ExtRst.Rphssub.St.CchCharacters[:]) * 2
		for i := uint16(0); i <= rgchDataSize; i++ {
			s.ExtRst.Rphssub.St.RgchData = append(s.ExtRst.Rphssub.St.RgchData, stream[iOft(&oft, 0):iOft(&oft, 2)]...)
		}

		//The number of elements in this array is rphssub.crun
		phRunsSizeL := helpers.BytesToUint16(s.ExtRst.Rphssub.Crun[:])
		for i := uint16(0); i <= phRunsSizeL; i++ {
			var phRuns PhRuns
			copy(phRuns.IchFirst[:], stream[iOft(&oft, 0):iOft(&oft, 2)])
			copy(phRuns.IchMom[:], stream[iOft(&oft, 0):iOft(&oft, 2)])
			copy(phRuns.CchMom[:], stream[iOft(&oft, 0):iOft(&oft, 2)])

			s.ExtRst.Rgphruns = append(s.ExtRst.Rgphruns, phRuns)
		}
	}
}

func iOft(offset *uint32, inc uint32) uint32 {
	*offset = *offset + inc
	return *offset
}

func (s *XLUnicodeRichExtendedString) String() string {

	if s.FHighByte&1 == 1 {
		name := helpers.BytesToUints16(s.Rgb[:])
		runes := utf16.Decode(name)
		return string(runes)
	} else {

		return string(s.Rgb[:])
	}

}
