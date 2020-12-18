package record

import (
	"bytes"
	"fmt"
	"github.com/metakeule/fmtdate"
	"github.com/shakinm/xlsReader/helpers"
	"github.com/shakinm/xlsReader/xls/structure"
	"strconv"
	"strings"
)

//FORMAT: Number Format

var FormatRecord = []byte{0x1E, 0x04} //(41Eh)

/*
The FORMAT record describes a number format in the workbook.
All the FORMAT records should appear together in a BIFF file. The order of FORMAT
records in an existing BIFF file should not be changed. It is possible to write custom
number formats in a file, but they should be added at the end of the existing FORMAT
records.

Record Data
Offset		Field Name		Size		Contents
------------------------------------------------
4			ifmt			2			Format index code (for internal use only)
6			cch				2			Length of the string
7			grbit			1			Option Flags (described in Unicode Strings in BIFF8 section)
8			rgb				var			Array of string characters

Excel uses the ifmt structure to identify built-in formats when it reads a file that was
created by a different localized version. For more information about built-in formats,
see "XF".

*/

type Format struct {
	ifmt     [2]byte
	cch      [2]byte
	grbit    [1]byte
	rgb      []byte
	vers     []byte
	stFormat structure.XLUnicodeRichExtendedString
}

func (r *Format) Read(stream []byte, vers []byte) {

	r.vers = vers

	if bytes.Compare(vers, FlagBIFF8) == 0 {
		copy(r.ifmt[:], stream[0:2])
		_ = r.stFormat.Read(stream[2:])
	} else {
		copy(r.ifmt[:], stream[:2])
		copy(r.cch[:], stream[2:4])
		r.rgb = make([]byte, helpers.BytesToUint16(r.cch[:]))
		copy(r.rgb[:], stream[4:])
	}

}

func (r *Format) String() string {

	if bytes.Compare(r.vers, FlagBIFF8) == 0 {
		return r.stFormat.String()
	}
	strLen := helpers.BytesToUint16(r.cch[:])
	return strings.TrimSpace(string(decodeWindows1251(bytes.Trim(r.rgb[:int(strLen)], "\x00"))))

}

func (r *Format) GetIndex() int {
	return int(helpers.BytesToUint16(r.ifmt[:]))
}

func (r *Format) GetFormatString(data structure.CellData) string {
	if r.GetIndex() >= 164 {

		if data.GetType() == "*record.LabelSSt" {
			return data.GetString()
		}
		if data.GetType() == "*record.Label" {
			return data.GetString()
		}

		if data.GetType() == "*record.FakeBlank" {
			return data.GetString()
		}

		if data.GetType() == "*record.Blank" {
			return data.GetString()
		}

		if data.GetType() == "*record.BoolErr" {
			return data.GetString()
		}

		if data.GetType() == "*record.Number" || data.GetType() == "*record.Rk" {
			if r.String() == "General" || r.String() == "@" {
				return strconv.FormatFloat(data.GetFloat64(), 'f', -1, 64)
			} else if strings.Contains(r.String(), "%") {
				return fmt.Sprintf("%.2f", data.GetFloat64()*100) + "%"
			} else if strings.Contains(r.String(), "#") || strings.Contains(r.String(), ".00") {
				return fmt.Sprintf("%.2f", data.GetFloat64())
			} else if strings.Contains(r.String(), "0") {
				return fmt.Sprintf("%.f", data.GetFloat64())
			} else {
				t := helpers.TimeFromExcelTime(data.GetFloat64(), false)
				dateFormat := strings.ReplaceAll(r.String(), "HH:MM:SS", "hh:mm:ss")
				dateFormat = strings.ReplaceAll(dateFormat, "\\", "")
				return fmtdate.Format(dateFormat, t)
			}

		}

		return data.GetString()

	} else {
		if data.GetType() == "*record.Number" {
			return strconv.FormatFloat(data.GetFloat64(), 'f', -1, 64)
		}
	}
	return data.GetString()
}
