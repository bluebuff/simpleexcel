package simpleexcel

import (
	"github.com/xuri/excelize/v2"
)

type StyleNameAlias string

const (
	Title            StyleNameAlias = "title"
	Head             StyleNameAlias = "head"
	Text             StyleNameAlias = "text"
	Number           StyleNameAlias = "number"
	Decimals         StyleNameAlias = "decimals"
	SubtotalText     StyleNameAlias = "subtotal_text"
	SubtotalNumber   StyleNameAlias = "subtotal_number"
	SubtotalDecimals StyleNameAlias = "subtotal_decimals"

	NumberCondition           StyleNameAlias = "number_condition"
	DecimalsCondition         StyleNameAlias = "decimals_condition"
	SubtotalNumberCondition   StyleNameAlias = "subtotal_number_condition"
	SubtotalDecimalsCondition StyleNameAlias = "subtotal_decimals_condition"
)

type CallSetStyleFunc func(file *excelize.File) (int, error)

type StyleManager interface {
	Configure(opts ...func(StyleManager)) StyleManager
	Store(alias StyleNameAlias, style interface{}) int
	Get(alias StyleNameAlias) (int, bool)
}

type styleManager struct {
	styleIdMap map[StyleNameAlias]int
	file       *excelize.File
}

func NewStyleManager(file *excelize.File, opts ...func(StyleManager)) StyleManager {
	mng := &styleManager{styleIdMap: make(map[StyleNameAlias]int), file: file}
	return mng.Configure(opts...)
}

func (mng *styleManager) Configure(opts ...func(StyleManager)) StyleManager {
	for _, item := range opts {
		item(mng)
	}
	return mng
}

func (mng *styleManager) Store(alias StyleNameAlias, item interface{}) int {
	if id, ok := mng.styleIdMap[alias]; ok {
		return id
	}
	switch v := item.(type) {
	case int:
		mng.styleIdMap[alias] = v
	case string:
		styleId, _ := NewStyleFunc(v)(mng.file)
		mng.styleIdMap[alias] = styleId
	case CallSetStyleFunc:
		styleId, _ := v(mng.file)
		mng.styleIdMap[alias] = styleId
	}
	return mng.styleIdMap[alias]
}

func (mng *styleManager) Get(alias StyleNameAlias) (int, bool) {
	styleId, ok := mng.styleIdMap[alias]
	return styleId, ok
}

func NewStyleFunc(style string) CallSetStyleFunc {
	return func(file *excelize.File) (int, error) {
		return file.NewStyle(style)
	}
}
