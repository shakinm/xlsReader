package cfb

import (
	"bytes"
	"encoding/binary"
	"io"
	"os"
	"path/filepath"

	"github.com/shakinm/xlsReader/helpers"
)

// Cfb - Compound File Binary
type Cfb struct {
	header           Header
	file             io.ReadSeeker
	fLink            *os.File
	difatPositions   []uint32
	miniFatPositions []uint32
	dirs             []*Directory
}

// EntrySize - Directory array entry length
var EntrySize = 128

// DefaultDIFATEntries -Number FAT locations in DIFAT
var DefaultDIFATEntries = uint32(109)

// GetDirs - Get a list of directories
func (cfb *Cfb) GetDirs() []*Directory {
	return cfb.dirs
}

func (cfb *Cfb) CloseFile() error {
	return cfb.fLink.Close()
}

// OpenFile - Open document from the file
func OpenFile(filename string) (cfb Cfb, err error) {

	cfb.fLink, err = os.Open(filepath.Clean(filename))

	if err != nil {
		return cfb, err
	}

	cfb.file = cfb.fLink

	err = open(&cfb)

	return cfb, err
}

// OpenReader - Open document from the reader
func OpenReader(reader io.ReadSeeker) (cfb Cfb, err error) {

	cfb.file = reader

	if err != nil {
		return
	}

	err = open(&cfb)

	return
}

func open(cfb *Cfb) (err error) {

	err = cfb.getHeader()

	if err != nil {
		return err
	}

	err = cfb.getMiniFATSectors()

	if err != nil {
		return err
	}

	err = cfb.getFatSectors()

	if err != nil {
		return err
	}

	err = cfb.getDirectories()

	return err
}

func (cfb *Cfb) getHeader() (err error) {

	var bHeader = make([]byte, 4096)

	_, err = cfb.file.Read(bHeader)

	if err != nil {
		return
	}

	err = binary.Read(bytes.NewBuffer(bHeader), binary.LittleEndian, &cfb.header)

	if err != nil {
		return
	}

	err = cfb.header.validate()

	return
}

func (cfb *Cfb) getDirectories() (err error) {

	stream, err := cfb.getDataFromFatChain(helpers.BytesToUint32(cfb.header.FirstDirectorySectorLocation[:]))

	if err != nil {
		return err
	}
	var section = make([]byte, 0)

	for _, value := range stream {
		section = append(section, value)
		if len(section) == EntrySize {
			var dir Directory
			err = binary.Read(bytes.NewBuffer(section), binary.LittleEndian, &dir)
			if err == nil && dir.ObjectType != 0x00 {
				cfb.dirs = append(cfb.dirs, &dir)
			}

			section = make([]byte, 0)
		}

	}

	return

}

func (cfb *Cfb) getMiniFATSectors() (err error) {

	var section = make([]byte, 0)

	position := cfb.calculateOffset(cfb.header.FirstMiniFATSectorLocation[:])

	for i := uint32(0); i < helpers.BytesToUint32(cfb.header.NumberMiniFATSectors[:]); i++ {
		sector := NewSector(&cfb.header)
		err := cfb.getData(position, &sector.Data)

		if err != nil {
			return err
		}

		for _, value := range sector.getMiniFatFATSectorLocations() {
			section = append(section, value)
			if len(section) == 4 {
				cfb.miniFatPositions = append(cfb.miniFatPositions, helpers.BytesToUint32(section))
				section = make([]byte, 0)
			}
		}
		position = position + sector.SectorSize
	}

	return
}

func (cfb *Cfb) getFatSectors() (err error) { // nolint: gocyclo

	entries := DefaultDIFATEntries

	if helpers.BytesToUint32(cfb.header.NumberFATSectors[:]) < DefaultDIFATEntries {
		entries = helpers.BytesToUint32(cfb.header.NumberFATSectors[:])
	}

	for i := uint32(0); i < entries; i++ {

		position := cfb.calculateOffset(cfb.header.getDIFATEntry(i))
		sector := NewSector(&cfb.header)

		err := cfb.getData(position, &sector.Data)

		if err != nil {
			return err
		}

		cfb.difatPositions = append(cfb.difatPositions, sector.values(EntrySize)...)

	}

	if bytes.Compare(cfb.header.FirstDIFATSectorLocation[:], ENDOFCHAIN) == 0 {
		return
	}

	position := cfb.calculateOffset(cfb.header.FirstDIFATSectorLocation[:])
	var section = make([]byte, 0)
	for i := uint32(0); i < helpers.BytesToUint32(cfb.header.NumberDIFATSectors[:]); i++ {
		sector := NewSector(&cfb.header)
		err := cfb.getData(position, &sector.Data)

		if err != nil {
			return err
		}

		for _, value := range sector.getFATSectorLocations() {
			section = append(section, value)
			if len(section) == 4 {

				position = cfb.calculateOffset(section)
				sectorF := NewSector(&cfb.header)
				err := cfb.getData(position, &sectorF.Data)

				if err != nil {
					return err
				}
				cfb.difatPositions = append(cfb.difatPositions, sectorF.values(EntrySize)...)

				section = make([]byte, 0)
			}

		}

		position = cfb.calculateOffset(sector.getNextDIFATSectorLocation())

	}

	return
}
func (cfb *Cfb) getDataFromMiniFat(miniFatSectorLocation uint32, offset uint32) (data []byte, err error) {

	point := cfb.calculateMiniFatOffset(offset)

	containerStreamBytes, _ := cfb.getDataFromFatChain(miniFatSectorLocation)
	containerStream := bytes.NewReader(containerStreamBytes)

	for {

		sector := NewMiniFatSector(&cfb.header)

		_, err := containerStream.ReadAt(sector.Data, int64(point))

		if err != nil {
			return data, err
		}

		data = append(data, sector.Data...)

		if cfb.miniFatPositions[offset] == helpers.BytesToUint32(ENDOFCHAIN) {
			break
		}

		offset = cfb.miniFatPositions[offset]

		point = cfb.calculateMiniFatOffset(offset)

	}

	return data, err
}

func (cfb *Cfb) getDataFromFatChain(offset uint32) (data []byte, err error) {

	for {
		sector := NewSector(&cfb.header)
		point := cfb.sectorOffset(offset)

		err = cfb.getData(point, &sector.Data)

		if err != nil {
			return data, err
		}

		data = append(data, sector.Data...)
		offset = cfb.difatPositions[offset]
		if offset == helpers.BytesToUint32(ENDOFCHAIN) {
			break
		}
	}

	return data, err
}

// OpenObject - Get object stream
func (cfb *Cfb) OpenObject(object *Directory, root *Directory) (reader io.ReadSeeker, err error) {

	if helpers.BytesToUint32(object.StreamSize[:]) < uint32(helpers.BytesToUint16(cfb.header.MiniStreamCutoffSize[:])) {

		data, err := cfb.getDataFromMiniFat(root.GetStartingSectorLocation(), object.GetStartingSectorLocation())

		if err != nil {
			return reader, err
		}

		reader = bytes.NewReader(data)
	} else {

		data, err := cfb.getDataFromFatChain(object.GetStartingSectorLocation())

		if err != nil {
			return reader, err
		}

		reader = bytes.NewReader(data)

	}

	return reader, err
}

func (cfb *Cfb) getData(offset uint32, data *[]byte) (err error) {

	_, err = cfb.file.Seek(int64(offset), 0)

	if err != nil {
		return
	}

	_, err = cfb.file.Read(*data)

	if err != nil {
		return
	}
	return

}

func (cfb *Cfb) sectorOffset(sid uint32) uint32 {
	return (sid + 1) * cfb.header.sectorSize()
}

func (cfb *Cfb) calculateMiniFatOffset(sid uint32) (n uint32) {

	return sid * 64
}

func (cfb *Cfb) calculateOffset(sectorID []byte) (n uint32) {

	if len(sectorID) == 4 {
		n = helpers.BytesToUint32(sectorID)
	}
	if len(sectorID) == 2 {
		n = uint32(binary.LittleEndian.Uint16(sectorID))
	}
	return (n * cfb.header.sectorSize()) + cfb.header.sectorSize()
}
