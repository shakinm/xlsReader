package xls

// Record struct
type Record struct {
	recordNumber [2]byte
	recordDataLength [2]byte
	recordData []byte
}
