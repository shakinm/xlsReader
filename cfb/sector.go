package cfb

import (
	"bytes"
	"encoding/binary"
)

// FREESECT - Specifies an unallocated sector in the FAT, Mini FAT, or DIFAT
var FREESECT = []byte{0xFF, 0xFF, 0xFF, 0xFF}

// ENDOFCHAIN -  End of a linked chain of sectors.
var ENDOFCHAIN = []byte{0xFE, 0xFF, 0xFF, 0xFF}

// FATSECT - Specifies a FAT sector in the FAT.
var FATSECT = []byte{0xFD, 0xFF, 0xFF, 0xFF}

// DIFSECT - Specifies a DIFAT sector in the FAT.
var DIFSECT = []byte{0xFC, 0xFF, 0xFF, 0xFF}

// Sector struct
type Sector struct {
	SectorSize uint32
	Data       []byte
}

func (s *Sector) getSector() *Sector {
	return s
}

func (s *Sector) findBlock(block []byte) bool {

	var section = make([]byte, 0)
	for _, value := range s.Data {
		section = append(section, value)
		if len(section) == 4 {
			if bytes.Compare(section, block) == 0 {
				return true
			}
			section = make([]byte, 0)

		}
	}
	return false
}

func (s *Sector) getFATSectorLocations() []byte {
	return s.Data[0 : s.SectorSize-4]
}

func (s *Sector) getMiniFatFATSectorLocations() []byte {
	return s.Data[0 : s.SectorSize]
}

func (s *Sector) getNextDIFATSectorLocation() []byte {
	return s.Data[s.SectorSize-4:]
}

// NewSector - Create new Sector struct fot FAT
func NewSector(header *Header) Sector {
	return Sector{
		SectorSize: header.sectorSize(),
		Data:       make([]byte, header.sectorSize()),
	}

}

// NewMiniFatSector - Create new Sector struct for MiniFat
func NewMiniFatSector(header *Header) Sector {
	return Sector{
		SectorSize: 64,
		Data:       make([]byte, 64),
	}
}


func (s *Sector) values(length int) (res []uint32 ) {

	  res = make([]uint32, length)

	buf := bytes.NewBuffer(s.Data)

	 _ = binary.Read(buf, binary.LittleEndian, res) // nolint: gosec

	return res
}