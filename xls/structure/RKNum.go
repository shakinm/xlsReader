package structure

import (
	"math"
	"strconv"

	"github.com/shakinm/xlsReader/helpers"
)

type RKNum [4]byte

func (r *RKNum) number() (intNum int64, floatNum float64, isFloat bool) {
	// Fix from: https://github.com/shakinm/xlsReader/issues/35
	rk := helpers.BytesToUint32(r[:])
	isFloat = rk&0x02 == 0
	isMul := rk&0x01 == 1
	if isFloat {
		floatNum = math.Float64frombits(uint64(rk&0xfffffffc) << 32)
		if isMul {
			floatNum /= 100
		}
	} else {
		intNum32 := uint32(rk >> 2)
		if rk&0x80000000 > 0 {
			intNum32 = intNum32 | 0xC0000000
		}
		if isMul {
			intNum32 /= 100
		}
		intNum = int64(int32(intNum32))
	}
	return
}

func (r *RKNum) GetFloat() (fn float64) {
	i, f, isFloat := r.number()
	if isFloat {
		fn = f
	} else {
		fn = float64(i)
	}
	return fn
}

func (r *RKNum) GetInt64() (in int64) {
	i, _, isFloat := r.number()
	if !isFloat {
		in = i
	}
	return in
}

func (r *RKNum) GetString() (s string) {
	i, f, isFloat := r.number()
	if isFloat {
		return strconv.FormatFloat(f, 'f', -1, 64)
	}
	return strconv.FormatInt(i, 10)
}
