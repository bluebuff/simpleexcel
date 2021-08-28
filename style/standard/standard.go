package standard

import (
	"github.com/bluebuff/simpleexcel/v2/style"
)

const (
	titleStyle                   = `{"fill":{"type":"pattern","pattern":18,"color":["#E8E8E8"]},"font":{"bold":true,"italic":false,"family":"正楷","size":18,"color":"#000000"},"alignment":{"horizontal":"center","vertical":"center"},"border":[{"type":"left","color":"#000000","style":1},{"type":"top","color":"#000000","style":1},{"type":"right","color":"#000000","style":1},{"type":"bottom","color":"#000000","style":1}]}`
	headStyle                    = `{"fill":{"type":"pattern","pattern":1,"color":["#E8E8E8"]},"font":{"bold":true,"italic":false,"family":"正楷","size":13,"color":"#000000"},"alignment":{"horizontal":"center","vertical":"center"},"border":[{"type":"left","color":"#000000","style":1},{"type":"top","color":"#000000","style":1},{"type":"right","color":"#000000","style":1},{"type":"bottom","color":"#000000","style":1}]}`
	textStyle                    = `{"font":{"bold":false,"italic":false,"family":"正楷","size":12,"color":"#000000"},"alignment":{"horizontal":"center","vertical":"center"},"border":[{"type":"left","color":"#000000","style":1},{"type":"top","color":"#000000","style":1},{"type":"right","color":"#000000","style":1},{"type":"bottom","color":"#000000","style":1}]}`
	numberStyle                  = `{"font":{"bold":false,"italic":false,"family":"Times New Roman","size":12,"color":"#000000"},"alignment":{"horizontal":"right","vertical":"center"},"border":[{"type":"left","color":"#000000","style":1},{"type":"top","color":"#000000","style":1},{"type":"right","color":"#000000","style":1},{"type":"bottom","color":"#000000","style":1}]}`
	numberConditionStyle         = `{"font":{"bold":false,"italic":false,"family":"Times New Roman","size":12,"color":"#FF0000"},"alignment":{"horizontal":"right","vertical":"center"},"border":[{"type":"left","color":"#000000","style":1},{"type":"top","color":"#000000","style":1},{"type":"right","color":"#000000","style":1},{"type":"bottom","color":"#000000","style":1}]}`
	decimalsStyle                = `{"number_format":4,"font":{"bold":false,"italic":false,"family":"Times New Roman","size":12,"color":"#000000"},"alignment":{"horizontal":"right","vertical":"center"},"border":[{"type":"left","color":"#000000","style":1},{"type":"top","color":"#000000","style":1},{"type":"right","color":"#000000","style":1},{"type":"bottom","color":"#000000","style":1}]}`
	decimalsConditionStyle       = `{"number_format":4,"font":{"bold":false,"italic":false,"family":"Times New Roman","size":12,"color":"#FF0000"},"alignment":{"horizontal":"right","vertical":"center"},"border":[{"type":"left","color":"#000000","style":1},{"type":"top","color":"#000000","style":1},{"type":"right","color":"#000000","style":1},{"type":"bottom","color":"#000000","style":1}]}`
	subtotalTextStyle            = `{"fill":{"type":"pattern","pattern":1,"color":["#E8E8E8"]},"font":{"bold":false,"italic":false,"family":"正楷","size":13,"color":"#000000"},"alignment":{"horizontal":"center","vertical":"center"},"border":[{"type":"left","color":"#000000","style":1},{"type":"top","color":"#000000","style":1},{"type":"right","color":"#000000","style":1},{"type":"bottom","color":"#000000","style":1}]}`
	subtotalNumberStyle          = `{"fill":{"type":"pattern","pattern":1,"color":["#E8E8E8"]},"font":{"bold":false,"italic":false,"family":"Times New Roman","size":13,"color":"#000000"},"alignment":{"horizontal":"right","vertical":"center"},"border":[{"type":"left","color":"#000000","style":1},{"type":"top","color":"#000000","style":1},{"type":"right","color":"#000000","style":1},{"type":"bottom","color":"#000000","style":1}]}`
	subtotalNumberConditionStyle = `{"fill":{"type":"pattern","pattern":1,"color":["#E8E8E8"]},"font":{"bold":false,"italic":false,"family":"Times New Roman","size":13,"color":"#FF0000"},"alignment":{"horizontal":"right","vertical":"center"},"border":[{"type":"left","color":"#000000","style":1},{"type":"top","color":"#000000","style":1},{"type":"right","color":"#000000","style":1},{"type":"bottom","color":"#000000","style":1}]}`
	subtotalDecimalsStyle        = `{"fill":{"type":"pattern","pattern":1,"color":["#E8E8E8"]},"number_format":4,"font":{"bold":false,"italic":false,"family":"Times New Roman","size":13,"color":"#000000"},"alignment":{"horizontal":"right","vertical":"center"},"border":[{"type":"left","color":"#000000","style":1},{"type":"top","color":"#000000","style":1},{"type":"right","color":"#000000","style":1},{"type":"bottom","color":"#000000","style":1}]}`
	subtotalDecimalsCondition    = `{"fill":{"type":"pattern","pattern":1,"color":["#E8E8E8"]},"number_format":4,"font":{"bold":false,"italic":false,"family":"Times New Roman","size":13,"color":"#FF0000"},"alignment":{"horizontal":"right","vertical":"center"},"border":[{"type":"left","color":"#000000","style":1},{"type":"top","color":"#000000","style":1},{"type":"right","color":"#000000","style":1},{"type":"bottom","color":"#000000","style":1}]}`
)

func Style(styleMng style.StyleManager) {
	// 用于标题 例如明细总表大标题
	styleMng.Store(style.Title, style.NewStyleFunc(titleStyle))
	// 用于头部
	styleMng.Store(style.Head, style.NewStyleFunc(headStyle))
	// 用于普通文本
	styleMng.Store(style.Text, style.NewStyleFunc(textStyle))
	// 用于数字
	styleMng.Store(style.Number, style.NewStyleFunc(numberStyle))
	// 当符合条件时，数字使用条件样式
	styleMng.Store(style.NumberCondition, style.NewStyleFunc(numberConditionStyle))
	// 用于小数
	styleMng.Store(style.Decimals, style.NewStyleFunc(decimalsStyle))
	// 用于小计（组合单位小计）
	styleMng.Store(style.SubtotalText, style.NewStyleFunc(subtotalTextStyle))
	// 当符合条件时，小数使用条件样式
	styleMng.Store(style.DecimalsCondition, style.NewStyleFunc(decimalsConditionStyle))
	// 用于小计数量
	styleMng.Store(style.SubtotalNumber, style.NewStyleFunc(subtotalNumberStyle))
	// 当符合条件时，小计金额使用条件样式
	styleMng.Store(style.SubtotalNumberCondition, style.NewStyleFunc(subtotalNumberConditionStyle))
	// 用于小计金额
	styleMng.Store(style.SubtotalDecimals, style.NewStyleFunc(subtotalDecimalsStyle))
	// 当符合条件时，小计金额使用条件样式
	styleMng.Store(style.SubtotalDecimalsCondition, style.NewStyleFunc(subtotalDecimalsCondition))
}
