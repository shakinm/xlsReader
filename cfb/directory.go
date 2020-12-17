package cfb

import (
	"encoding/binary"
	"github.com/Alliera/xlsReader/helpers"
	"unicode/utf16"
)

// Directory - Compound File Directory Entry
type Directory struct {
	DirectoryEntryName       [64]byte
	DirectoryEntryNameLength [2]byte
	ObjectType               byte
	ColorFlag                [1]byte
	LeftSiblingID            [4]byte
	RightSiblingID           [4]byte
	ChildID                  [4]byte
	CLSID                    [16]byte
	StateBits                [4]byte
	CreationTime             [8]byte
	ModifiedTime             [8]byte
	StartingSectorLocation   [4]byte
	StreamSize               [8]byte
}

//Name - Directory Name
func (d *Directory) Name() string {

	size := binary.LittleEndian.Uint16(d.DirectoryEntryNameLength[:])
	if size > 0 {
		size = size - 1
	}
	name := helpers.BytesToUints16(d.DirectoryEntryName[:size])
	runes := utf16.Decode(name)
	return string(runes)
}

//GetStartingSectorLocation - The start sector of the object
func (d *Directory) GetStartingSectorLocation() uint32 {

	return helpers.BytesToUint32(d.StartingSectorLocation[:])
}

//GetStreamSize - Object size
func (d *Directory) GetStreamSize() uint32 {

	return helpers.BytesToUint32(d.StreamSize[:])
}
