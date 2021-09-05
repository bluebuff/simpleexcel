package option

import "github.com/shopspring/decimal"

type compareSymbol string

const (
	LT compareSymbol = "<"
	LE compareSymbol = "<="
	EQ compareSymbol = "=="
	NE compareSymbol = "!="
	GE compareSymbol = ">="
	GT compareSymbol = ">"
)

// int32
func cmpInt32(a, b int32, c compareSymbol) (r bool) {
	switch {
	case c == LT && a < b:
		r = true
	case c == LE && a <= b:
		r = true
	case c == EQ && a == b:
		r = true
	case c == NE && a != b:
		r = true
	case c == GE && a >= b:
		r = true
	case c == GT && a > b:
		r = true
	}
	return
}

// int64
func cmpInt64(a, b int64, c compareSymbol) (r bool) {
	switch {
	case c == LT && a < b:
		r = true
	case c == LE && a <= b:
		r = true
	case c == EQ && a == b:
		r = true
	case c == NE && a != b:
		r = true
	case c == GE && a >= b:
		r = true
	case c == GT && a > b:
		r = true
	}
	return
}

// uint32
func cmpUint32(a, b uint32, c compareSymbol) (r bool) {
	switch {
	case c == LT && a < b:
		r = true
	case c == LE && a <= b:
		r = true
	case c == EQ && a == b:
		r = true
	case c == NE && a != b:
		r = true
	case c == GE && a >= b:
		r = true
	case c == GT && a > b:
		r = true
	}
	return
}

// uint64
func cmpUint64(a, b uint64, c compareSymbol) (r bool) {
	switch {
	case c == LT && a < b:
		r = true
	case c == LE && a <= b:
		r = true
	case c == EQ && a == b:
		r = true
	case c == NE && a != b:
		r = true
	case c == GE && a >= b:
		r = true
	case c == GT && a > b:
		r = true
	}
	return
}

func cmpFloat32(a, b float32, c compareSymbol) bool {
	fa := decimal.NewFromFloat32(a)
	fb := decimal.NewFromFloat32(b)
	return cmp(fa, fb, c)
}

func cmpFloat64(a, b float64, c compareSymbol) bool {
	fa := decimal.NewFromFloat(a)
	fb := decimal.NewFromFloat(b)
	return cmp(fa, fb, c)
}

func cmp(a, b decimal.Decimal, c compareSymbol) (r bool) {
	switch {
	case c == LT && a.LessThan(b):
		r = true
	case c == LE && a.LessThanOrEqual(b):
		r = true
	case c == EQ && a.Equal(b):
		r = true
	case c == NE && !a.Equal(b):
		r = true
	case c == GE && a.GreaterThanOrEqual(b):
		r = true
	case c == GT && a.GreaterThan(b):
		r = true
	}
	return
}
