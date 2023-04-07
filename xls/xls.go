package xls

import (
	"encoding/binary"
	"github.com/shakinm/xlsReader/cfb"
	"io"
)

// OpenFile - Open document from the file
func OpenFile(fileName string) (workbook Workbook, err error) {

	adaptor, err := cfb.OpenFile(fileName)

	defer adaptor.CloseFile()

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
	dir := adaptor.GetDirs()
	if len(dir) == 0 {
		return workbook, errors.New("directory not found")
	}

	var book *cfb.Directory
	root := dir[0]
	for _, file := range dir {
		switch file.Name() {
		case "Workbook":
			if book == nil {
				book = file
			}
		case "Book":
			book = file
		case "Root Entry":
			root = file
		}
	}
	if book == nil {
		return workbook, errors.New("workbook not found")
	}

	size := binary.LittleEndian.Uint32(book.StreamSize[:])
	reader, err := adaptor.OpenObject(book, root)
	if err != nil {
		return workbook, err
	}

	return readStream(reader, size)
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
