package xls

import (
	"encoding/binary"
	"github.com/shakinm/xlsReader/cfb"
	"io"
)

// OpenFile - Open document from the file
func OpenFile(fileName string) (workbook Workbook, err error) {

	adaptor, err := cfb.OpenFile(fileName)

	if err != nil {
		return workbook, err
	}
	return openCfb(adaptor)
}

// OpenReader - Open document from the file reader
func OpenReader(fileReader io.ReadSeeker) (workbook Workbook, err error) {

	adaptor, err := cfb.OpenReader(fileReader)

	if err != nil {
		return workbook, err
	}
	return openCfb(adaptor)
}

// OpenFile - Open document from the file
func openCfb(adaptor cfb.Cfb) (workbook Workbook, err error) {
	var book *cfb.Directory
	var root *cfb.Directory
	for _, dir := range adaptor.GetDirs() {
		fn := dir.Name()

		if fn == "Workbook" {
			if book == nil {
				book = dir
			}
		}
		if fn == "Book" {
			book = dir

		}
		if fn == "Root Entry" {
			root = dir
		}

	}

	if book != nil {
		size := binary.LittleEndian.Uint32(book.StreamSize[:])

		reader, err := adaptor.OpenObject(book, root)

		if err != nil {
			return workbook, err
		}

		return readStream(reader, size)

	}

	return workbook, err
}

func readStream(reader io.ReadSeeker, streamSize uint32) (workbook Workbook, err error) {

	stream := make([]byte, streamSize)

	_, err = reader.Read(stream)

	if err != nil {
		return workbook, nil
	}


	if err != nil {
		return workbook, nil
	}

	err = workbook.read(stream)

	if err != nil {
		return workbook, nil
	}

	for k := range workbook.sheets {
		sheet, err := workbook.GetSheet(k)

		if err != nil {
			return workbook, nil
		}

		err = sheet.read(stream)

		if err != nil {
			return workbook, nil
		}
	}

	return
}
