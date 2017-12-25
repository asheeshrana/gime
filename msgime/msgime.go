package msgime

import (
	"fmt"
	"io"
	"os"
)

type fieldPosition struct {
	offset int
	size   int
}

var uuidMimeTypeMap = map[string]string{
	"00020906-0000-0000-c000-000000000046": "application/msword",
	"00020820-0000-0000-c000-000000000046": "application/vnd.ms-excel",
	"00020810-0000-0000-c000-000000000046": "application/vnd.ms-excel",
}

var headerMap = map[string]fieldPosition{
	"FileIdentifier":                                 fieldPosition{offset: 0, size: 8},
	"UUIDOfFile":                                     fieldPosition{offset: 8, size: 16},
	"RevisionNumber":                                 fieldPosition{offset: 24, size: 2},
	"VersionNumber":                                  fieldPosition{offset: 26, size: 2},
	"ByteOrderIdentifier":                            fieldPosition{offset: 28, size: 2},
	"SizeOfSector":                                   fieldPosition{offset: 30, size: 2},
	"SizeOfShortSector":                              fieldPosition{offset: 32, size: 2},
	"Reserved":                                       fieldPosition{offset: 34, size: 10},
	"TotalSectors":                                   fieldPosition{offset: 44, size: 4},
	"FirstSectorID":                                  fieldPosition{offset: 48, size: 4},
	"Reserved1":                                      fieldPosition{offset: 52, size: 4},
	"MinSizeOfStdStream":                             fieldPosition{offset: 56, size: 4},
	"FirstShortSectorID":                             fieldPosition{offset: 60, size: 4},
	"TotalSectorsUsedForShortSectorAllocationTable":  fieldPosition{offset: 64, size: 4},
	"FistMasterSectorID":                             fieldPosition{offset: 68, size: 4},
	"TotalSectorsUsedForMasterSectorAllocationTable": fieldPosition{offset: 72, size: 4},
	"FirstPartOfMasterAllocationTable":               fieldPosition{offset: 76, size: 436},
}

var directoryMap = map[string]fieldPosition{
	"EntryName":                   fieldPosition{offset: 0, size: 64},
	"SizeOfEntryNameInCharacters": fieldPosition{offset: 64, size: 2},
	"Type":                        fieldPosition{offset: 66, size: 1},  //00H = Empty 03H = LockBytes (unknown), 01H = User storage 04H = Property (unknown), 02H = User stream 05H = Root storage
	"NodeColorr":                  fieldPosition{offset: 67, size: 1},  //00H = Red 01H = Black. It is a read-black tree
	"DirIDOfLeftChild":            fieldPosition{offset: 68, size: 4},  //DirID of the left child node inside the red-black tree of all direct members of the parent storage (if this entry is a user storage or stream), –1 if there is no left child
	"DirIDOfRighttChild":          fieldPosition{offset: 72, size: 4},  //DirID of the right child node inside the red-black tree of all direct members of the parent storage (if this entry is a user storage or stream), –1 if there is no right child
	"DirIDOfRoot":                 fieldPosition{offset: 76, size: 4},  //DirID of the root node entry of the red-black tree of all storage members (if this entry is a storage), –1 otherwise
	"CLSID":                       fieldPosition{offset: 80, size: 16}, //UUID representing CLSID
	"UserFlags":                   fieldPosition{offset: 96, size: 4},
	"EntryCreationTimpestamp":     fieldPosition{offset: 100, size: 8},
	"EntryModificationTimpestamp": fieldPosition{offset: 108, size: 8},
	"FistSectorID":                fieldPosition{offset: 116, size: 4},
	"TotalStreamSizeInBytes":      fieldPosition{offset: 120, size: 4},
	"Reserved":                    fieldPosition{offset: 124, size: 4},
}

type CompoundFile interface {
	GetMimeType() string
	PrintFileInfo()
}

type defaultCompoundFileInterface interface {
	CompoundFile
	//Private methods
	getValueFromHeader(fieldname string) []byte
	getValueFromRootDirectory(fieldname string) []byte
	isLittleEndian() bool
	setHeader(header []byte) CompoundFile
	setRootDirectory(rootDirectory []byte) CompoundFile
	setFilename(filepath string) CompoundFile
}

type defaultCompoundFile struct {
	filename           string
	header             []byte
	rootDirectoryEntry []byte
}

func (cFile *defaultCompoundFile) GetMimeType() string {
	clsID := cFile.getValueFromRootDirectory("CLSID")
	uuID := decodeValueAsUUID(cFile.isLittleEndian(), clsID)
	if mimeType, ok := uuidMimeTypeMap[uuID]; ok {
		return mimeType
	}
	return "application/octet-stream"
}

func (cFile *defaultCompoundFile) PrintFileInfo() {
	printValue("FileIdentifier", cFile.getValueFromHeader("FileIdentifier"))
	fmt.Printf("Filename = %s\n", cFile.filename)
	fmt.Printf("UUIDOfFile = %s\n", decodeValueAsUUID(cFile.isLittleEndian(), cFile.getValueFromHeader("UUIDOfFile")))
	printValue("RevisionNumber", cFile.getValueFromHeader("RevisionNumber"))
	printValue("VersionNumber", cFile.getValueFromHeader("VersionNumber"))
	fmt.Printf("LittleEndian = %t", cFile.isLittleEndian())
}

func (cFile *defaultCompoundFile) getValueFromHeader(fieldname string) []byte {
	var fieldValue []byte
	if fieldInfo, ok := headerMap[fieldname]; ok {
		fieldValue = cFile.header[fieldInfo.offset : fieldInfo.offset+fieldInfo.size]
	}
	return fieldValue
}

func (cFile *defaultCompoundFile) getValueFromRootDirectory(fieldname string) []byte {
	var fieldValue []byte
	if fieldInfo, ok := directoryMap[fieldname]; ok {
		fieldValue = cFile.rootDirectoryEntry[fieldInfo.offset : fieldInfo.offset+fieldInfo.size]
	}
	return fieldValue
}

func (cFile *defaultCompoundFile) isLittleEndian() bool {
	byteOrder := cFile.getValueFromHeader("ByteOrderIdentifier")
	return byteOrder[0] == 0xFE
}

func (cFile *defaultCompoundFile) setHeader(header []byte) CompoundFile {
	cFile.header = header
	return cFile
}
func (cFile *defaultCompoundFile) setFilename(filepath string) CompoundFile {
	cFile.filename = filepath
	return cFile
}
func (cFile *defaultCompoundFile) setRootDirectory(rootDirectory []byte) CompoundFile {
	cFile.rootDirectoryEntry = rootDirectory
	return cFile
}

func NewCompoundFile(filepath string) (CompoundFile, error) {
	var err error
	var cfile defaultCompoundFileInterface = &defaultCompoundFile{filename: filepath}
	file, _ := os.Open(filepath)

	defer file.Close()
	//Header always starts at offset 0 and is of size 512
	cfile.setHeader(read(file, 0, 512))
	littleEndian := cfile.isLittleEndian()
	sectorID := decodeValueAsUInt16(littleEndian, cfile.getValueFromHeader("FirstSectorID"))
	sectorSize := decodeValueAsUInt16(littleEndian, cfile.getValueFromHeader("SizeOfSector"))
	sectorPosition := getSectorPosition(sectorID, sectorSize)

	//Sector is always of size 128
	cfile.setRootDirectory(read(file, int64(sectorPosition), 128))

	return cfile, err
}

func read(file *os.File, offset int64, size int) []byte {
	var buffer = make([]byte, size)
	file.Seek(offset, io.SeekStart)
	io.ReadFull(file, buffer)
	return buffer
}

func getFileSize(filepath string) (int64, error) {
	file, err := os.Open(filepath)
	if err != nil {
		panic(err)
	}
	fs, err := file.Stat()
	if err != nil {
		// Could not obtain stat, handle error
		return -1, err
	}

	return fs.Size(), nil
}

func decodeValueAsUUID(littleEndian bool, value []byte) string {
	//Microsoft uses mixed endian https://en.wikipedia.org/wiki/Universally_unique_identifier
	//So we will ignore the flag and decode first 3 components as little endian and last 2 components as big endian
	var bytes1To4 = decodeValueAsUInt64(true, value[0:4])
	var bytes5To6 = decodeValueAsUInt64(true, value[4:6])
	var bytes7To8 = decodeValueAsUInt64(true, value[6:8])
	var bytes9To10 = decodeValueAsUInt64(false, value[8:10])
	var bytes11To16 = decodeValueAsUInt64(false, value[10:16])

	return fmt.Sprintf("%08x-%04x-%04x-%04x-%012x", bytes1To4, bytes5To6, bytes7To8, bytes9To10, bytes11To16)
}

func decodeValueAsUInt64(littleEndian bool, value []byte) uint64 {
	//Not using binary.littleendian.Uint16 because it expects the value to be 8 byte only
	var returnValue uint64
	for i := 0; i < len(value); i++ {
		if littleEndian {
			returnValue = (returnValue << 8) | uint64(value[len(value)-(i+1)])
		} else {
			returnValue = (returnValue << 8) | uint64(value[i])
		}
	}
	return returnValue
}

func decodeValueAsUInt16(littleEndian bool, value []byte) uint16 {
	//Not using binary.littleendian.Uint16 because it expects the value to be 2 byte only
	var returnValue uint16
	for i := 0; i < len(value); i++ {
		if littleEndian {
			returnValue = (returnValue << 8) | uint16(value[len(value)-(i+1)])
		} else {
			returnValue = (returnValue << 8) | uint16(value[i])
		}
	}
	return returnValue
}

func decodeValueAsByteArray(littleEndian bool, value []byte) []byte {
	var returnValue = value
	if littleEndian {
		returnValue = make([]byte, len(value))
		copy(returnValue, value)
		for index, byteValue := range value {
			returnValue[len(value)-(index+1)] = byteValue
		}
	}
	return returnValue
}

func printValue(fieldname string, value []byte) {
	fmt.Printf("%s = ", fieldname)
	for _, byteValue := range value {
		fmt.Printf("%02x ", byteValue)
	}
	fmt.Println()
}

func getSectorPosition(sectorID uint16, sectorSize uint16) uint64 {
	return 512 + uint64(sectorID)*calcPower(2, sectorSize)
}

func calcPower(x uint16, y uint16) uint64 {
	//Not using golang math as it returns float and don't want to even deal with possibilites of precision issues due to using float instead of an int
	if y == 0 {
		return 1
	}
	var result uint64 = 1
	var multiplier = uint64(x)
	for i := y; i > 1; {
		if y%2 == 0 {
			multiplier = multiplier * multiplier
			i = i / 2
		} else {
			result = result * multiplier
			i = i - 1
		}
	}
	return result * multiplier
}
