package structure

type CellData interface {
	GetString() string
	GetFloat64() float64
	GetInt64() int64
	GetXFIndex() int
	GetType() string
}
