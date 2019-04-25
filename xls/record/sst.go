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
	i         int
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

func (s *SST) Read(readType string, grbit byte, prevLen int32) () {

	defer r()

	if len(s.RgbSrc) == 0 {
		return
	}

	oft := uint32(0)

	for {

		var _rgb structure.XLUnicodeRichExtendedString
		var rgbSize uint16

		//offset 2
		copy(_rgb.Cch[:], s.RgbSrc[iOft(&oft, 0):iOft(&oft, 2)])
		//offset 2
		_rgb.FHighByte = s.RgbSrc[iOft(&oft, 0):iOft(&oft, 1)][0]

		if _rgb.FHighByte&1 == 1 {
			rgbSize = helpers.BytesToUint16(_rgb.Cch[:]) * 2
			if readType == "continue" && prevLen > 0 && grbit == 0 {
				rgbSize = rgbSize - (rgbSize-uint16(prevLen-3))/2

			}

		} else {
			rgbSize = helpers.BytesToUint16(_rgb.Cch[:])
			if readType == "continue" && prevLen > 0 && grbit == 1 {
				rgbSize = uint16(prevLen-3) + (rgbSize-uint16(prevLen-3))*2
			}
		}
		readType = ""

		if _rgb.FHighByte>>3&1 == 1 { // if fRichSt == 1

			copy(_rgb.CRun[:], s.RgbSrc[iOft(&oft, 0):iOft(&oft, 2)])
		}

		if _rgb.FHighByte>>2&1 == 1 { //fExtSt  == 1
			//offset 4
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
			for i := uint16(0); i <= phRunsSizeL; i++ {
				var phRuns structure.PhRuns
				copy(phRuns.IchFirst[:], s.RgbSrc[iOft(&oft, 0):iOft(&oft, 2)])
				copy(phRuns.IchMom[:], s.RgbSrc[iOft(&oft, 0):iOft(&oft, 2)])
				copy(phRuns.CchMom[:], s.RgbSrc[iOft(&oft, 0):iOft(&oft, 2)])

				_rgb.ExtRst.Rgphruns = append(_rgb.ExtRst.Rgphruns, phRuns)
			}
		}

		if len(s.RgbSrc) >= int(oft) {

			s.Rgb = append(s.Rgb, _rgb)

			s.RgbSrc = s.RgbSrc[int(oft):]
			oft = 0
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
