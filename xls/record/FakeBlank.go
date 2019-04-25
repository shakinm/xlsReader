package record

import "reflect"

//Fake record

type FakeBlank struct {
}

func (r *FakeBlank) GetString() (str string) {
	return str
}

func (r *FakeBlank) GetFloat64() (fl float64) {
	return fl
}
func (r *FakeBlank) GetInt64() (in int64) {
	return in
}

func (r *FakeBlank) GetType() string {
	return reflect.TypeOf(r).String()
}

func (r *FakeBlank) GetXFIndex() int {
	return 15 //The last record ( ixfe=15 ) is the default cell XF for the workbook.
}
