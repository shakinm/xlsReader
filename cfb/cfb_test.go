package cfb

import (
	"fmt"
	"testing"
)

func TestOpenFile(t *testing.T) {
	ole, err := OpenFile("bigfile.xls")
	fmt.Println(err)
	fmt.Println(ole.dirs)
}
