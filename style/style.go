package style

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

type StyleExpress string

type CallSetStyleFunc func(file *excelize.File) (int, error)

var styleMap map[StyleNameAlias]CallSetStyleFunc

func Register(alias StyleNameAlias, styleFunc CallSetStyleFunc) {
	styleMap[alias] = styleFunc
}

func GetStyleFunc(alias StyleNameAlias) (CallSetStyleFunc, bool) {
	styleFunc, ok := styleMap[alias]
	return styleFunc, ok
}

func NewStyle(style string) CallSetStyleFunc {
	return func(file *excelize.File) (int, error) {
		return file.NewStyle(style)
	}
}

func NewConditionStyle(style string) CallSetStyleFunc {
	return func(file *excelize.File) (int, error) {
		return file.NewConditionalStyle(style)
	}
}

type StyleManager struct {
	styleIdMap map[StyleNameAlias]int
}

func NewStyleManager(file *excelize.File) *StyleManager {
	styleIdMap := make(map[StyleNameAlias]int)
	for k, fc := range styleMap {
		styleId, _ := fc(file)
		styleIdMap[k] = styleId
	}
	return &StyleManager{
		styleIdMap: styleIdMap,
	}
}

func (mng *StyleManager) Store(alias StyleNameAlias, styleId int) bool {
	if _, ok := mng.styleIdMap[alias]; ok {
		return false
	}
	mng.styleIdMap[alias] = styleId
	return true
}

func (mng *StyleManager) Get(alias StyleNameAlias) (int, bool) {
	styleId, ok := mng.styleIdMap[alias]
	return styleId, ok
}