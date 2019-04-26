package structure

import (
	"github.com/shakinm/xlsReader/helpers"
	"math"
	"strconv"
)

type RKNum [4]byte

func (r *RKNum) number() (intNum int64, floatNum float64, isFloat bool) {
	rk := helpers.BytesToUint32(r[:])

	val := uint64(rk >> 2)
	rkType := uint(rk << 30 >> 30)

	var fn float64
	switch rkType {
	case 0:
		fn = math.Float64frombits(uint64(rk&0xfffffffc) << 32)
		isFloat = true
	case 1:

		fn = math.Float64frombits(uint64(rk&0xfffffffc)<<32) / 100
		isFloat = true
	case 3:
		fn = float64(val) / 100
		isFloat = true
	}

	return int64(val), float64(fn), isFloat
}

func (r *RKNum) GetFloat() (fn float64) {
	i, f, isFloat := r.number()
	if isFloat {
		fn = f
	} else {
		fn=float64(i)
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
