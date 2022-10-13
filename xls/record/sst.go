package record

import (
	"github.com/shakinm/xlsReader/helpers"
	"github.com/shakinm/xlsReader/xls/structure"
	"io"
)

// SST: Shared String Table

var SSTRecord = [2]byte{0xFC, 0x00} //(FCh)

/*
The SST record contains string constants.


Record Data — BIFF8
Offset		Name		Size		Contents
--------------------------------------------
4 			cstTotal 	4 			Total number of strings in the shared string table and
									extended string table ( EXTSST record)
8 			cstUnique 	4 			Number of unique strings in the shared string table
12 			rgb 		var 		Array of unique unicode strings (XLUnicodeRichExtendedString).

*/

type SST struct {
	CstTotal  [4]byte
	CstUnique [4]byte
	RgbSrc    []byte
	Rgb       []structure.XLUnicodeRichExtendedString
	chLen     int
	ByteLen   int

	// These are needed to properly handle CONTINUE records.
	// CONTINUE record contains grbit in the first byte unless it's a formatting run
	// so we need to know whether all the string bytes have been consumed.
	// OpenOffice.org - Microsoft Excel File Format - section 5.21
	RgbDone bool
	Grbit   byte
}

func (s *SST) RgbAppend(bts []byte) (err error) {
	for _, value := range bts {
		s.RgbSrc = append(s.RgbSrc, value)
	}

	return err
}

func r() (err error) {
	if r := recover(); r != nil {
		return io.EOF
	}
	return
}

func abs(x int) int {
	if x < 0 {
		return -x
	}
	return x
}

func (s *SST) Read(readType string, grbit byte, prevLen int32) () {

	defer r()

	if len(s.RgbSrc) == 0 {
		return
	}
	oft := uint32(0)
	for {

		var _rgb structure.XLUnicodeRichExtendedString
		var rgbSize int

		cch := int(helpers.BytesToUint16(s.RgbSrc[0:2]))

		if readType != "continue" {
			grbit = s.RgbSrc[2:3][0]
		}

		if readType == "continue" && prevLen == 0 && s.ByteLen == 0 {
			grbit = s.RgbSrc[2:3][0]
		}

		readType = ""

		s.Grbit = grbit

		headLen := 3
		headLen += int(grbit>>2&1) * 4
		headLen += int(grbit>>3&1) * 2

		if cch >= (len(s.RgbSrc)-headLen)/(1+int(grbit&1)) || s.ByteLen > 0 {

			addBytesLen := (len(s.RgbSrc) - headLen) - s.ByteLen

			if cch-s.chLen > addBytesLen/(1+int(grbit&1)) {
				s.chLen = s.chLen + addBytesLen/(1+int(grbit&1))
				s.ByteLen = s.ByteLen + addBytesLen
				return
			} else {

				s.ByteLen = s.ByteLen + (cch-s.chLen)*(1+int(grbit&1))
				s.chLen = cch
				rgbSize = s.ByteLen
			}

		} else {
			rgbSize = cch * (1 + int(grbit&1))
		}

		copy(_rgb.Cch[:], s.RgbSrc[iOft(&oft, 0):iOft(&oft, 2)])

		_rgb.FHighByte = s.RgbSrc[iOft(&oft, 0):iOft(&oft, 1)][0]

		if _rgb.FHighByte>>3&1 == 1 { // if fRichSt == 1
			copy(_rgb.CRun[:], s.RgbSrc[iOft(&oft, 0):iOft(&oft, 2)])
		}

		if _rgb.FHighByte>>2&1 == 1 { //fExtSt  == 1
			copy(_rgb.CbExtRst[:], s.RgbSrc[iOft(&oft, 0):iOft(&oft, 4)])
		}

		//offset rgbSize
		_rgb.Rgb = make([]byte, uint32(rgbSize))
		copy(_rgb.Rgb[0:], s.RgbSrc[iOft(&oft, 0):iOft(&oft, uint32(rgbSize))])

		if _rgb.FHighByte>>3&1 == 1 { // if fRichSt == 1
			cRunSize := helpers.BytesToUint16(_rgb.CRun[:])
			for i := uint16(0); i <= cRunSize-1; i++ {
				var rgRun structure.FormatRun
				copy(rgRun.Ich[:], s.RgbSrc[iOft(&oft, 0):iOft(&oft, 2)])
				copy(rgRun.Ifnt.Ifnt[:], s.RgbSrc[iOft(&oft, 0):iOft(&oft, 2)])
				_rgb.RgRun = append(_rgb.RgRun, rgRun)
			}
		}

		if _rgb.FHighByte>>2&1 == 1 { //fExtSt  == 1
			copy(_rgb.ExtRst.Reserved[:], s.RgbSrc[iOft(&oft, 0):iOft(&oft, 2)])
			copy(_rgb.ExtRst.Cb[:], s.RgbSrc[iOft(&oft, 0):iOft(&oft, 2)])

			copy(_rgb.ExtRst.Phs.Ifnt.Ifnt[:], s.RgbSrc[iOft(&oft, 0):iOft(&oft, 2)])
			copy(_rgb.ExtRst.Phs.Info[:], s.RgbSrc[iOft(&oft, 0):iOft(&oft, 2)])

			copy(_rgb.ExtRst.Rphssub.Crun[:], s.RgbSrc[iOft(&oft, 0):iOft(&oft, 2)])
			copy(_rgb.ExtRst.Rphssub.Cch[:], s.RgbSrc[iOft(&oft, 0):iOft(&oft, 2)])

			copy(_rgb.ExtRst.Rphssub.St.CchCharacters[:], s.RgbSrc[iOft(&oft, 0):iOft(&oft, 2)])

			rgchDataSize := helpers.BytesToUint16(_rgb.ExtRst.Rphssub.St.CchCharacters[:]) * 2
			for i := uint16(0); i <= rgchDataSize; i++ {
				_rgb.ExtRst.Rphssub.St.RgchData = append(_rgb.ExtRst.Rphssub.St.RgchData, s.RgbSrc[iOft(&oft, 0):iOft(&oft, 2)]...)
			}

			//The number of elements in this array is rphssub.crun
			phRunsSizeL := helpers.BytesToUint16(_rgb.ExtRst.Rphssub.Crun[:])
			if phRunsSizeL > 0 {
				for i := uint16(0); i <= phRunsSizeL; i++ {
					var phRuns structure.PhRuns
					copy(phRuns.IchFirst[:], s.RgbSrc[iOft(&oft, 0):iOft(&oft, 2)])
					copy(phRuns.IchMom[:], s.RgbSrc[iOft(&oft, 0):iOft(&oft, 2)])
					copy(phRuns.CchMom[:], s.RgbSrc[iOft(&oft, 0):iOft(&oft, 2)])

					_rgb.ExtRst.Rgphruns = append(_rgb.ExtRst.Rgphruns, phRuns)
				}
			}
		}

		if len(s.RgbSrc) >= int(oft) {
			s.Rgb = append(s.Rgb, _rgb)
			s.RgbSrc = s.RgbSrc[int(oft):]
			s.chLen = 0
			s.ByteLen = 0
			oft = 0

			if len(s.RgbSrc) == 0 {
				return
			}

		} else {
			break
		}

	}

}

func iOft(offset *uint32, inc uint32) uint32 {
	*offset = *offset + inc
	return *offset
}

func (s *SST) NewSST(buf []byte) {
	copy(s.CstTotal[:], buf[:4])
	copy(s.CstUnique[:], buf[4:8])
	s.RgbSrc = append(s.RgbSrc, buf[8:]...)
}
