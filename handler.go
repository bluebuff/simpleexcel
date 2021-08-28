package simpleexcel

import (
	"github.com/bluebuff/simpleexcel/v2/internal"
	"github.com/bluebuff/simpleexcel/v2/model"
)

type Handler func(styleMng StyleManager, cell *model.Cell)

// compare
func CompareLessAndStyleInt32(compareValue int32) Handler {
	return func(styleMng StyleManager, cell *model.Cell) {
		if v, ok := cell.Value.(int32); ok && v < compareValue {
			cell.StyleID, _ = styleMng.Get(NumberCondition)
		}
	}
}

func CompareLessAndStyleInt64(compareValue int64) Handler {
	return func(styleMng StyleManager, cell *model.Cell) {
		if v, ok := cell.Value.(int64); ok && v < compareValue {
			cell.StyleID, _ = styleMng.Get(NumberCondition)
		}
	}
}

func CompareLessAndStyleUint32(compareValue uint32) Handler {
	return func(styleMng StyleManager, cell *model.Cell) {
		if v, ok := cell.Value.(uint32); ok && v < compareValue {
			cell.StyleID, _ = styleMng.Get(NumberCondition)
		}
	}
}

func CompareLessAndStyleUint64(compareValue uint64) Handler {
	return func(styleMng StyleManager, cell *model.Cell) {
		if v, ok := cell.Value.(uint64); ok && v < compareValue {
			cell.StyleID, _ = styleMng.Get(NumberCondition)
		}
	}
}

func CompareLessAndStyleFloat32(compareValue float32) Handler {
	return func(styleMng StyleManager, cell *model.Cell) {
		if v, ok := cell.Value.(float32); ok && v < compareValue {
			cell.StyleID, _ = styleMng.Get(DecimalsCondition)
		}
	}
}

func CompareLessAndStyleFloat64(compareValue float64) Handler {
	return func(styleMng StyleManager, cell *model.Cell) {
		if v, ok := cell.Value.(float64); ok && v < compareValue {
			cell.StyleID, _ = styleMng.Get(DecimalsCondition)
		}
	}
}

// custom style
func CustomStyleByAlias(alias StyleNameAlias) Handler {
	return func(styleMng StyleManager, cell *model.Cell) {
		cell.StyleID, _ = styleMng.Get(alias)
	}
}

// custom style
func CustomStyleJson(json string) Handler {
	return func(styleMng StyleManager, cell *model.Cell) {
		md5 := internal.Md5(json)
		cell.StyleID = styleMng.Store(StyleNameAlias(md5), json)
	}
}
