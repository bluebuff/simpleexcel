package option

import (
	"github.com/bluebuff/simpleexcel/v2/internal"
	"github.com/bluebuff/simpleexcel/v2/model"
	"github.com/bluebuff/simpleexcel/v2/style"
	"strings"
)

type Option func(styleMng style.StyleManager, cell *model.Cell)

// compare int32
func CompareAndStyleInt32(symbol compareSymbol, compareValue int32) Option {
	return func(styleMng style.StyleManager, cell *model.Cell) {
		if v, ok := cell.Value.(int32); ok && cmpInt32(v, compareValue, symbol) {
			cell.StyleID, _ = styleMng.Get(style.NumberCondition)
		}
	}
}

// compare int64
func CompareAndStyleInt64(symbol compareSymbol, compareValue int64) Option {
	return func(styleMng style.StyleManager, cell *model.Cell) {
		if v, ok := cell.Value.(int64); ok && cmpInt64(v, compareValue, symbol) {
			cell.StyleID, _ = styleMng.Get(style.NumberCondition)
		}
	}
}

// compare uint32
func CompareAndStyleUint32(symbol compareSymbol, compareValue uint32) Option {
	return func(styleMng style.StyleManager, cell *model.Cell) {
		if v, ok := cell.Value.(uint32); ok && cmpUint32(v, compareValue, symbol) {
			cell.StyleID, _ = styleMng.Get(style.NumberCondition)
		}
	}
}

// compare uint64
func CompareAndStyleUint64(symbol compareSymbol, compareValue uint64) Option {
	return func(styleMng style.StyleManager, cell *model.Cell) {
		if v, ok := cell.Value.(uint64); ok && cmpUint64(v, compareValue, symbol) {
			cell.StyleID, _ = styleMng.Get(style.NumberCondition)
		}
	}
}

// compare float32
func CompareAndStyleFloat32(symbol compareSymbol, compareValue float32) Option {
	return func(styleMng style.StyleManager, cell *model.Cell) {
		if v, ok := cell.Value.(float32); ok && cmpFloat32(v, compareValue, symbol) {
			cell.StyleID, _ = styleMng.Get(style.DecimalsCondition)
		}
	}
}

// compare float64
func CompareAndStyleFloat64(symbol compareSymbol, compareValue float64) Option {
	return func(styleMng style.StyleManager, cell *model.Cell) {
		if v, ok := cell.Value.(float64); ok && cmpFloat64(v, compareValue, symbol) {
			cell.StyleID, _ = styleMng.Get(style.DecimalsCondition)
		}
	}
}

// compare interface  in (int32,int64,uint32,uint64,float32,float64)
func CompareAndStyleInterface(symbol compareSymbol, compareValue interface{}) Option {
	var compareFunc Option
	switch v := compareValue.(type) {
	case int32:
		compareFunc = CompareAndStyleInt32(symbol, v)
	case int64:
		compareFunc = CompareAndStyleInt64(symbol, v)
	case uint32:
		compareFunc = CompareAndStyleUint32(symbol, v)
	case uint64:
		compareFunc = CompareAndStyleUint64(symbol, v)
	case float32:
		compareFunc = CompareAndStyleFloat32(symbol, v)
	case float64:
		compareFunc = CompareAndStyleFloat64(symbol, v)
	default:
		compareFunc = func(styleMng style.StyleManager, cell *model.Cell) {
			// empty
		}
	}
	return compareFunc
}

// compare swap | int32
func CompareAndSwapInt32(symbol compareSymbol, compareValue, swap int32) Option {
	return func(styleMng style.StyleManager, cell *model.Cell) {
		if v, ok := cell.Value.(int32); ok && cmpInt32(v, compareValue, symbol) {
			cell.Value = swap
		}
	}
}

// compare swap | int64
func CompareAndSwapInt64(symbol compareSymbol, compareValue, swap int64) Option {
	return func(styleMng style.StyleManager, cell *model.Cell) {
		if v, ok := cell.Value.(int64); ok && cmpInt64(v, compareValue, symbol) {
			cell.Value = swap
		}
	}
}

// compare swap | uint32
func CompareAndSwapUint32(symbol compareSymbol, compareValue, swap uint32) Option {
	return func(styleMng style.StyleManager, cell *model.Cell) {
		if v, ok := cell.Value.(uint32); ok && cmpUint32(v, compareValue, symbol) {
			cell.Value = swap
		}
	}
}

// compare swap | uint64
func CompareAndSwapUint64(symbol compareSymbol, compareValue, swap uint64) Option {
	return func(styleMng style.StyleManager, cell *model.Cell) {
		if v, ok := cell.Value.(uint64); ok && cmpUint64(v, compareValue, symbol) {
			cell.Value = swap
		}
	}
}

// compare swap | float32
func CompareAndSwapFloat32(symbol compareSymbol, compareValue, swap float32) Option {
	return func(styleMng style.StyleManager, cell *model.Cell) {
		if v, ok := cell.Value.(float32); ok && cmpFloat32(v, compareValue, symbol) {
			cell.Value = swap
		}
	}
}

// compare swap | float64
func CompareAndSwapFloat64(symbol compareSymbol, compareValue, swap float64) Option {
	return func(styleMng style.StyleManager, cell *model.Cell) {
		if v, ok := cell.Value.(float64); ok && cmpFloat64(v, compareValue, symbol) {
			cell.Value = swap
		}
	}
}

// empty | dft
func WhenEmptyAndSwapString(dft string) Option {
	return func(styleMng style.StyleManager, cell *model.Cell) {
		if v, ok := cell.Value.(string); ok && strings.TrimSpace(v) == "" {
			cell.Value = dft
		}
	}
}

// custom style
func CustomStyleByAlias(alias style.StyleNameAlias) Option {
	return func(styleMng style.StyleManager, cell *model.Cell) {
		cell.StyleID, _ = styleMng.Get(alias)
	}
}

// custom style
func CustomStyleJson(json string) Option {
	return func(styleMng style.StyleManager, cell *model.Cell) {
		md5 := internal.Md5(json)
		cell.StyleID = styleMng.Store(style.StyleNameAlias(md5), json)
	}
}
