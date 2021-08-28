package common

import (
	"fmt"
	"github.com/xuri/excelize/v2"
)

type StyleNameAlias string

const (
	Title                     StyleNameAlias = "title"
	Head                      StyleNameAlias = "head"
	Text                      StyleNameAlias = "text"
	Number                    StyleNameAlias = "number"
	NumberCondition           StyleNameAlias = "number_condition"
	Decimals                  StyleNameAlias = "decimals"
	DecimalsCondition         StyleNameAlias = "decimals_condition"
	SubtotalText              StyleNameAlias = "subtotal_text"
	SubtotalNumber            StyleNameAlias = "subtotal_number"
	SubtotalNumberCondition   StyleNameAlias = "subtotal_number_condition"
	SubtotalDecimals          StyleNameAlias = "subtotal_decimals"
	SubtotalDecimalsCondition StyleNameAlias = "subtotal_decimals_condition"
)

type callSetStyleFunc func(sheetName, hcell, vcell, expression string, compareVal float64)

type ConditionStyleGroup struct {
	StyleName    StyleNameAlias
	StyleExpress ConditionExpress
}

func LoadDefaultStyle(styleManager *StyleManager, file *excelize.File) {
	defaultStyleFuncList := GetDefaultRegisterStyle(file)
	for alias, style := range defaultStyleFuncList {
		styleManager.SetStyle(alias, style)
	}
}

type StyleManager struct {
	styleMap       map[StyleNameAlias]callSetStyleFunc
}

func NewStyleMng() *StyleManager {
	return &StyleManager{styleMap: make(map[StyleNameAlias]callSetStyleFunc)}
}

func (mng *StyleManager) SetStyle(alias StyleNameAlias, fun callSetStyleFunc) {
	mng.styleMap[alias] = fun
}

func (mng *StyleManager) GetStyleFunc(alias StyleNameAlias) (callSetStyleFunc, bool) {
	styleFunc, ok := mng.styleMap[alias]
	return styleFunc, ok
}

// 注册样式
func GetDefaultRegisterStyle(file *excelize.File) map[StyleNameAlias]callSetStyleFunc {
	styleMap := make(map[StyleNameAlias]callSetStyleFunc, 0)
	// 用于标题 例如明细总表大标题
	styleMap[Title] = newStyle(file, `{"font":{"bold":true,"italic":false,"family":"正楷","size":18,"color":"#000000"},"alignment":{"horizontal":"center","vertical":"center"},"border":[{"type":"left","color":"#000000","style":1},{"type":"top","color":"#000000","style":1},{"type":"right","color":"#000000","style":1},{"type":"bottom","color":"#000000","style":1}]}`)
	// 用于头部
	styleMap[Head] = newStyle(file, `{"fill":{"type":"pattern","pattern":1,"color":["#E8E8E8"]},"font":{"bold":true,"italic":false,"family":"正楷","size":12,"color":"#000000"},"alignment":{"horizontal":"center","vertical":"center"},"border":[{"type":"left","color":"#000000","style":1},{"type":"top","color":"#000000","style":1},{"type":"right","color":"#000000","style":1},{"type":"bottom","color":"#000000","style":1}]}`)
	// 用于普通文本
	styleMap[Text] = newStyle(file, `{"font":{"bold":false,"italic":false,"family":"正楷","size":12,"color":"#000000"},"alignment":{"horizontal":"center","vertical":"center"},"border":[{"type":"left","color":"#000000","style":1},{"type":"top","color":"#000000","style":1},{"type":"right","color":"#000000","style":1},{"type":"bottom","color":"#000000","style":1}]}`)
	// 用于数字
	styleMap[Number] = newStyle(file, `{"font":{"bold":false,"italic":false,"family":"Times New Roman","size":12,"color":"#000000"},"alignment":{"horizontal":"right","vertical":"center"},"border":[{"type":"left","color":"#000000","style":1},{"type":"top","color":"#000000","style":1},{"type":"right","color":"#000000","style":1},{"type":"bottom","color":"#000000","style":1}]}`)
	// 当符合条件时，数字使用条件样式
	styleMap[NumberCondition] = newConditionStyle(file, `{"alignment":{"horizontal":"right","vertical":"center"},"font":{"family":"Times New Roman","bold":false,"size":12,"color":"#FF0000"},"border":[{"type":"left","color":"#000000","style":1},{"type":"top","color":"#000000","style":1},{"type":"right","color":"#000000","style":1},{"type":"bottom","color":"#000000","style":1}]}`)
	// 用于小数
	styleMap[Decimals] = newStyle(file, `{"number_format":4,"font":{"bold":false,"italic":false,"family":"Times New Roman","size":12,"color":"#000000"},"alignment":{"horizontal":"right","vertical":"center"},"border":[{"type":"left","color":"#000000","style":1},{"type":"top","color":"#000000","style":1},{"type":"right","color":"#000000","style":1},{"type":"bottom","color":"#000000","style":1}]}`)
	// 用于小计（组合单位小计）
	styleMap[SubtotalText] = newStyle(file, `{"fill":{"type":"pattern","pattern":1,"color":["#E8E8E8"]},"font":{"bold":false,"italic":false,"family":"正楷","size":12,"color":"#000000"},"alignment":{"horizontal":"center","vertical":"center"},"border":[{"type":"left","color":"#000000","style":1},{"type":"top","color":"#000000","style":1},{"type":"right","color":"#000000","style":1},{"type":"bottom","color":"#000000","style":1}]}`)
	// 当符合条件时，小数使用条件样式
	styleMap[DecimalsCondition] = newConditionStyle(file, `{"number_format":4,"font":{"bold":false,"italic":false,"family":"Times New Roman","size":12,"color":"#FF0000"},"alignment":{"horizontal":"right","vertical":"center"},"border":[{"type":"left","color":"#000000","style":1},{"type":"top","color":"#000000","style":1},{"type":"right","color":"#000000","style":1},{"type":"bottom","color":"#000000","style":1}]}`)
	// 用于小计数量
	styleMap[SubtotalNumber] = newStyle(file, `{"fill":{"type":"pattern","pattern":1,"color":["#E8E8E8"]},"font":{"bold":false,"italic":false,"family":"Times New Roman","size":12,"color":"#000000"},"alignment":{"horizontal":"right","vertical":"center"},"border":[{"type":"left","color":"#000000","style":1},{"type":"top","color":"#000000","style":1},{"type":"right","color":"#000000","style":1},{"type":"bottom","color":"#000000","style":1}]}`)
	// 当符合条件时，小计金额使用条件样式
	styleMap[SubtotalNumberCondition] = newConditionStyle(file, `{"fill":{"type":"pattern","pattern":1,"color":["#E8E8E8"]},"alignment":{"horizontal":"right","vertical":"center"},"font":{"family":"Times New Roman","bold":false,"size":12,"color":"#FF0000"},"border":[{"type":"left","color":"#000000","style":1},{"type":"top","color":"#000000","style":1},{"type":"right","color":"#000000","style":1},{"type":"bottom","color":"#000000","style":1}]}`)
	// 用于小计金额
	styleMap[SubtotalDecimals] = newStyle(file, `{"fill":{"type":"pattern","pattern":1,"color":["#E8E8E8"]},"number_format":4,"font":{"bold":false,"italic":false,"family":"Times New Roman","size":12,"color":"#000000"},"alignment":{"horizontal":"right","vertical":"center"},"border":[{"type":"left","color":"#000000","style":1},{"type":"top","color":"#000000","style":1},{"type":"right","color":"#000000","style":1},{"type":"bottom","color":"#000000","style":1}]}`)
	// 当符合条件时，小计金额使用条件样式
	styleMap[SubtotalDecimalsCondition] = newConditionStyle(file, `{"fill":{"type":"pattern","pattern":1,"color":["#E8E8E8"]},"number_format":4,"alignment":{"horizontal":"right","vertical":"center"},"font":{"family":"Times New Roman","bold":false,"size":12,"color":"#FF0000"},"border":[{"type":"left","color":"#000000","style":1},{"type":"top","color":"#000000","style":1},{"type":"right","color":"#000000","style":1},{"type":"bottom","color":"#000000","style":1}]}`)
	return styleMap
}

func newStyle(file *excelize.File, style string) callSetStyleFunc {
	index, _ := file.NewStyle(style)
	return func(sheetName, hcell, vcell, expression string, compareVal float64) {
		file.SetCellStyle(sheetName, hcell, vcell, index)
	}
}

func newConditionStyle(file *excelize.File, style string) callSetStyleFunc {
	index, _ := file.NewConditionalStyle(style)
	return func(sheetName, hcell, vcell, expression string, compareVal float64) {
		area := fmt.Sprintf("%s:%s", hcell, vcell)
		formatSet := fmt.Sprintf(expression, index, compareVal)
		file.SetConditionalFormat(sheetName, area, formatSet)
	}
}
