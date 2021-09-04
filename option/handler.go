package option

import (
	"github.com/bluebuff/simpleexcel/v2/internal"
	"github.com/bluebuff/simpleexcel/v2/model"
	"github.com/bluebuff/simpleexcel/v2/style"
)

type Option func(styleMng style.StyleManager, cell *model.Cell)

// compare int32
// symbols: default <
func CompareAndStyleInt32(compareValue int32, symbols ...compareSymbol) Option {
	return func(styleMng style.StyleManager, cell *model.Cell) {
		if v, ok := cell.Value.(int32); ok && cmpInt32(v, compareValue, getSymbol(symbols...)) {
			cell.StyleID, _ = styleMng.Get(style.NumberCondition)
		}
	}
}

// compare int64
// symbols: default <
func CompareAndStyleInt64(compareValue int64, symbols ...compareSymbol) Option {
	return func(styleMng style.StyleManager, cell *model.Cell) {
		if v, ok := cell.Value.(int64); ok && cmpInt64(v, compareValue, getSymbol(symbols...)) {
			cell.StyleID, _ = styleMng.Get(style.NumberCondition)
		}
	}
}

// compare uint32
// symbols: default <
func CompareAndStyleUint32(compareValue uint32, symbols ...compareSymbol) Option {
	return func(styleMng style.StyleManager, cell *model.Cell) {
		if v, ok := cell.Value.(uint32); ok && cmpUint32(v, compareValue, getSymbol(symbols...)) {
			cell.StyleID, _ = styleMng.Get(style.NumberCondition)
		}
	}
}

// compare uint64
// symbols: default <
func CompareAndStyleUint64(compareValue uint64, symbols ...compareSymbol) Option {
	return func(styleMng style.StyleManager, cell *model.Cell) {
		if v, ok := cell.Value.(uint64); ok && cmpUint64(v, compareValue, getSymbol(symbols...)) {
			cell.StyleID, _ = styleMng.Get(style.NumberCondition)
		}
	}
}

// compare float32
// symbols: default <
func CompareAndStyleFloat32(compareValue float32, symbols ...compareSymbol) Option {
	return func(styleMng style.StyleManager, cell *model.Cell) {
		if v, ok := cell.Value.(float32); ok && cmpFloat32(v, compareValue, getSymbol(symbols...)) {
			cell.StyleID, _ = styleMng.Get(style.DecimalsCondition)
		}
	}
}

// compare float64
// symbols: default <
func CompareAndStyleFloat64(compareValue float64, symbols ...compareSymbol) Option {
	return func(styleMng style.StyleManager, cell *model.Cell) {
		if v, ok := cell.Value.(float64); ok && cmpFloat64(v, compareValue, getSymbol(symbols...)) {
			cell.StyleID, _ = styleMng.Get(style.DecimalsCondition)
		}
	}
}

// compare interface  in (int32,int64,uint32,uint64,float32,float64)
// symbols: default <
func CompareValueAndStyleInt32(compareValue interface{}, symbols ...compareSymbol) Option {
	var compareFunc Option
	switch v := compareValue.(type) {
	case int32:
		compareFunc = CompareAndStyleInt32(v, symbols...)
	case int64:
		compareFunc = CompareAndStyleInt64(v, symbols...)
	case uint32:
		compareFunc = CompareAndStyleUint32(v, symbols...)
	case uint64:
		compareFunc = CompareAndStyleUint64(v, symbols...)
	case float32:
		compareFunc = CompareAndStyleFloat32(v, symbols...)
	case float64:
		compareFunc = CompareAndStyleFloat64(v, symbols...)
	default:
		compareFunc = func(styleMng style.StyleManager, cell *model.Cell) {
			// empty
		}
	}
	return compareFunc
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
