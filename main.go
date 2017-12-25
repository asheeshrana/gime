package main

import (
	"fmt"

	"github.com/asheeshrana/gime/msgime"
)

func main() {
	var cfile, _ = msgime.NewCompoundFile("/temp/CV.doc")
	var mimeType = cfile.GetMimeType()
	fmt.Println("Mime type of the file = " + mimeType)
	cfile.PrintFileInfo()
}
