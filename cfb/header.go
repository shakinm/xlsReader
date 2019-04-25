package cfb

import (
	"bytes"
	"errors"
	"github.com/shakinm/xlsReader/helpers"
)

//HeaderSignature Identification signature for the compound file structure, and MUST be
//set to the value ...
var HeaderSignature = []byte{0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1}

//MajorVersion3 - Version number sign fot version 3
var MajorVersion3 = []byte{0x03, 0x00}

//MajorVersion4 - Version number sign fot version 4
var MajorVersion4 = []byte{0x04, 0x00}

//MajorVersion -Version number for breaking changes. This field MUST be set to either
//0x0003 (version 3) or 0x0004 (version 4).
var MajorVersion = [][]byte{MajorVersion3, MajorVersion4}

//ByteOrder - This field MUST be set to 0xFFFE. This field is a byte order mark for all integer
//fields, specifying little-endian byte order.
var ByteOrder = []byte{0xFE, 0xFF}

//SectorShiftForMajorVersion3 - If Major Version is 3, the Sector Shift MUST be 0x0009, specifying a sector size of 512 bytes.
var SectorShiftForMajorVersion3 = []byte{0x09, 0x00}

//SectorShiftForMajorVersion4 - If Major Version is 4, the Sector Shift MUST be 0x000C, specifying a sector size of 4096 bytes.
var SectorShiftForMajorVersion4 = []byte{0x0C, 0x00}

//MiniSectorShift - This field MUST be set to 0x0009, or 0x000c, depending on the Major
//Version field. This field specifies the sector size of the compound file as a power of 2.
var MiniSectorShift = []byte{0x06, 0x00}

//Reserved - This field MUST be set to all zeroes.
var Reserved = []byte{0x00, 0x00, 0x00, 0x00, 0x00, 0x00}

//NumberDirectorySectorsForMajorVersion3 - If Major Version is 3, the Number of Directory Sectors MUST be zero.
var NumberDirectorySectorsForMajorVersion3 = []byte{0x00, 0x00, 0x00, 0x00}

//MiniStreamCutoffSize - This integer field MUST be set to 0x00001000. This field
//specifies the maximum size of a user-defined data stream that is allocated from the mini FAT
//and mini stream, and that cutoff is 4,096 bytes. Any user-defined data stream that is greater than
//or equal to this cutoff size must be allocated as normal sectors from the FAT.
var MiniStreamCutoffSize = []byte{0x00, 0x10, 0x00, 0x00}

// Header - The Compound File Header structure
type Header struct {
	HeaderSignature              [8]byte
	HeaderCLSID                  [16]byte
	MinorVersion                 [2]byte
	MajorVersion                 [2]byte
	ByteOrder                    [2]byte
	SectorShift                  [2]byte
	MiniSectorShift              [2]byte
	Reserved                     [6]byte
	NumberDirectorySectors       [4]byte
	NumberFATSectors             [4]byte
	FirstDirectorySectorLocation [4]byte
	TransactionSignatureNumber   [4]byte
	MiniStreamCutoffSize         [4]byte
	FirstMiniFATSectorLocation   [4]byte
	NumberMiniFATSectors         [4]byte
	FirstDIFATSectorLocation     [4]byte
	NumberDIFATSectors           [4]byte
	DIFAT                        [3584]byte
}

func (h *Header) getDIFATEntry(i uint32) []byte {
	return h.DIFAT[i*4:(i*4)+4]
}

func (h *Header) sectorSize() (size uint32) {
	if bytes.Compare(h.MajorVersion[:], MajorVersion3) == 0 {
		size = 512
	}
	if bytes.Compare(h.MajorVersion[:], MajorVersion4) == 0 {
		size = 4096
	}

	return size
}

func (h *Header) validate() (err error) { // nolint: gocyclo

	if bytes.Compare(h.HeaderSignature[:], HeaderSignature) != 0 {
		return errors.New(`Identification signature for the compound file structure, and MUST be set to the value 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1`)
	}

	if !helpers.BytesInSlice(h.MajorVersion[:], MajorVersion) {
		return errors.New(`Version number for breaking changes. This structure MUST be set to either 0x0003 (version 3) or 0x0004 (version 4)`)
	}

	if bytes.Compare(h.ByteOrder[:], ByteOrder) != 0 {
		return errors.New(`Byte Order MUST be set to 0xFFFE. This structure is a byte order mark for all integer fields, specifying little-endian byte order`)
	}

	if bytes.Compare(h.MajorVersion[:], MajorVersion3) == 0 && bytes.Compare(h.SectorShift[:], SectorShiftForMajorVersion3) != 0 {
		return errors.New(`If Major Version is 3, the Sector Shift MUST be 0x0009, specifying a sector size of 512 bytes`)
	}

	if bytes.Compare(h.MajorVersion[:], MajorVersion4) == 0 && bytes.Compare(h.SectorShift[:], SectorShiftForMajorVersion4) != 0 {
		return errors.New(`If Major Version is 4, the Sector Shift MUST be 0x000C, specifying a sector size of 4,096 bytes`)
	}

	if bytes.Compare(h.MiniSectorShift[:], MiniSectorShift) != 0 {
		return errors.New(`Mini Sector Shift MUST be set to 0x0006. This structure specifies the sector size of the Mini Stream as a power of 2. The sector size of the Mini Stream MUST be 64 bytes`)
	}

	if bytes.Compare(h.Reserved[:], Reserved) != 0 {
		return errors.New(`Reserved MUST be set to all zeroes`)
	}

	if bytes.Compare(h.MiniStreamCutoffSize[:], MiniStreamCutoffSize) != 0 {
		return errors.New(`Mini Stream Cutoff Size structure MUST be set to 0x00001000`)
	}

	if bytes.Compare(h.MajorVersion[:], MajorVersion3) == 0 && bytes.Compare(h.NumberDirectorySectors[:], NumberDirectorySectorsForMajorVersion3) != 0 {
		return errors.New(`if Major Version is 3, the Number of Directory Sectors MUST be zero`)
	}

	if bytes.Compare(h.MajorVersion[:], MajorVersion4) == 0 {
		for i := 513; i <= 4096; i++ {
			if h.DIFAT[i] != 0x00 {
				return errors.New(`For version 4 compound files, the header size (512 bytes) is less than the sector size (4,096 bytes), so the remaining part of the header (3,584 bytes) MUST be filled with all zeroes`)
			}
		}

	}

	return
}

