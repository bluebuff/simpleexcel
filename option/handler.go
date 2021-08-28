package option

import (
	"github.com/bluebuff/simpleexcel/v2/internal"
	"github.com/bluebuff/simpleexcel/v2/model"
	"github.com/bluebuff/simpleexcel/v2/style"
)

type Option func(styleMng style.StyleManager, cell *model.Cell)

// compare
func CompareLessAndStyleInt32(compareValue int32) Option {
	return func(styleMng style.StyleManager, cell *model.Cell) {
		if v, ok := cell.Value.(int32); ok && v < compareValue {
			cell.StyleID, _ = styleMng.Get(style.NumberCondition)
		}
	}
}

func CompareLessAndStyleInt64(compareValue int64) Option {
	return func(styleMng style.StyleManager, cell *model.Cell) {
		if v, ok := cell.Value.(int64); ok && v < compareValue {
			cell.StyleID, _ = styleMng.Get(style.NumberCondition)
		}
	}
}

func CompareLessAndStyleUint32(compareValue uint32) Option {
	return func(styleMng style.StyleManager, cell *model.Cell) {
		if v, ok := cell.Value.(uint32); ok && v < compareValue {
			cell.StyleID, _ = styleMng.Get(style.NumberCondition)
		}
	}
}

func CompareLessAndStyleUint64(compareValue uint64) Option {
	return func(styleMng style.StyleManager, cell *model.Cell) {
		if v, ok := cell.Value.(uint64); ok && v < compareValue {
			cell.StyleID, _ = styleMng.Get(style.NumberCondition)
		}
	}
}

func CompareLessAndStyleFloat32(compareValue float32) Option {
	return func(styleMng style.StyleManager, cell *model.Cell) {
		if v, ok := cell.Value.(float32); ok && v < compareValue {
			cell.StyleID, _ = styleMng.Get(style.DecimalsCondition)
		}
	}
}

func CompareLessAndStyleFloat64(compareValue float64) Option {
	return func(styleMng style.StyleManager, cell *model.Cell) {
		if v, ok := cell.Value.(float64); ok && v < compareValue {
			cell.StyleID, _ = styleMng.Get(style.DecimalsCondition)
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
