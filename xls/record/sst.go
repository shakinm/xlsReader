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

		if cch >= (len(s.RgbSrc)-3)/(1+int(grbit)) || s.ByteLen > 0 {

			addBytesLen := (len(s.RgbSrc) - 3) - s.ByteLen

			if cch-s.chLen > addBytesLen/(1+int(grbit)) {
				s.chLen = s.chLen + addBytesLen/(1+int(grbit))
				s.ByteLen = s.ByteLen + addBytesLen
				return
			} else {

				s.ByteLen = s.ByteLen + (cch-s.chLen)*(1+int(grbit))
				s.chLen = cch
				rgbSize = s.ByteLen + 3
			}

		} else {
			rgbSize = cch*(1+int(grbit)) + 3
		}

		_rgb.Read(s.RgbSrc[iOft(&oft, 0):iOft(&oft, uint32(rgbSize))])

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
