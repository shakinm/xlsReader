package record

import "github.com/shakinm/xlsReader/xls/structure"

//EXTSST: Extended Shared String Table

var ExtSstRecord = [2]byte{0xFF, 0x00} //(FFh)

type ExtSST struct {
	dsst      [2]byte
	rgisstinf []structure.ISSTINF
}

func (r *ExtSST) GetRgisstinf() []structure.ISSTINF {
	return r.rgisstinf
}

func (r *ExtSST) Read(stream []byte) {
	copy(r.dsst[:], stream[:2])

	for i := 0; i <= len(stream[2:])/6; i++ {
		sPoint := 2 + (i * 6)
		var inf structure.ISSTINF
		copy(inf.Cb[:], stream[sPoint:sPoint+4])
		copy(inf.Ib[:], stream[sPoint+4:sPoint+6])
		copy(inf.Reserved[:], stream[sPoint+6:sPoint+8])
		r.rgisstinf=append(r.rgisstinf, inf)
	}
}
