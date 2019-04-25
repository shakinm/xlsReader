package helpers

import (
	"bytes"
	"encoding/binary"
)

func BytesToUint64(b []byte) uint64 {
	return binary.LittleEndian.Uint64(b)
}

func BytesToUint32(b []byte) uint32 {
	return binary.LittleEndian.Uint32(b)
}

func BytesToUint16(b []byte) uint16 {
	return binary.LittleEndian.Uint16(b)
}

func BytesInSlice(a []byte, list [][]byte) bool {
	for _, b := range list {
		if bytes.Compare(a, b) == 0 {
			return true
		}
	}
	return false
}


func BytesToUints16(b []byte) (res []uint16) {

	var section = make([]byte, 0)
	for _, value := range b {
		section = append(section, value)
		if len(section) == 2 {
			res = append(res, binary.LittleEndian.Uint16(section))

			section = make([]byte, 0)
		}
	}
	return
}