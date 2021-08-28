package standard

import "github.com/bluebuff/simpleexcel/v2/style"

const (
	titleStyle                   = `{"font":{"bold":true,"italic":false,"family":"正楷","size":18,"color":"#000000"},"alignment":{"horizontal":"center","vertical":"center"},"border":[{"type":"left","color":"#000000","style":1},{"type":"top","color":"#000000","style":1},{"type":"right","color":"#000000","style":1},{"type":"bottom","color":"#000000","style":1}]}`
	headStyle                    = `{"fill":{"type":"pattern","pattern":1,"color":["#E8E8E8"]},"font":{"bold":true,"italic":false,"family":"正楷","size":12,"color":"#000000"},"alignment":{"horizontal":"center","vertical":"center"},"border":[{"type":"left","color":"#000000","style":1},{"type":"top","color":"#000000","style":1},{"type":"right","color":"#000000","style":1},{"type":"bottom","color":"#000000","style":1}]}`
	textStyle                    = `{"font":{"bold":false,"italic":false,"family":"正楷","size":12,"color":"#000000"},"alignment":{"horizontal":"center","vertical":"center"},"border":[{"type":"left","color":"#000000","style":1},{"type":"top","color":"#000000","style":1},{"type":"right","color":"#000000","style":1},{"type":"bottom","color":"#000000","style":1}]}`
	numberStyle                  = `{"font":{"bold":false,"italic":false,"family":"Times New Roman","size":12,"color":"#000000"},"alignment":{"horizontal":"right","vertical":"center"},"border":[{"type":"left","color":"#000000","style":1},{"type":"top","color":"#000000","style":1},{"type":"right","color":"#000000","style":1},{"type":"bottom","color":"#000000","style":1}]}`
	numberConditionStyle         = `{"alignment":{"horizontal":"right","vertical":"center"},"font":{"family":"Times New Roman","bold":false,"size":12,"color":"#FF0000"},"border":[{"type":"left","color":"#000000","style":1},{"type":"top","color":"#000000","style":1},{"type":"right","color":"#000000","style":1},{"type":"bottom","color":"#000000","style":1}]}`
	decimalsStyle                = `{"number_format":4,"font":{"bold":false,"italic":false,"family":"Times New Roman","size":12,"color":"#000000"},"alignment":{"horizontal":"right","vertical":"center"},"border":[{"type":"left","color":"#000000","style":1},{"type":"top","color":"#000000","style":1},{"type":"right","color":"#000000","style":1},{"type":"bottom","color":"#000000","style":1}]}`
	subtotalTextStyle            = `{"fill":{"type":"pattern","pattern":1,"color":["#E8E8E8"]},"font":{"bold":false,"italic":false,"family":"正楷","size":12,"color":"#000000"},"alignment":{"horizontal":"center","vertical":"center"},"border":[{"type":"left","color":"#000000","style":1},{"type":"top","color":"#000000","style":1},{"type":"right","color":"#000000","style":1},{"type":"bottom","color":"#000000","style":1}]}`
	decimalsConditionStyle       = `{"number_format":4,"font":{"bold":false,"italic":false,"family":"Times New Roman","size":12,"color":"#FF0000"},"alignment":{"horizontal":"right","vertical":"center"},"border":[{"type":"left","color":"#000000","style":1},{"type":"top","color":"#000000","style":1},{"type":"right","color":"#000000","style":1},{"type":"bottom","color":"#000000","style":1}]}`
	subtotalNumberStyle          = `{"fill":{"type":"pattern","pattern":1,"color":["#E8E8E8"]},"font":{"bold":false,"italic":false,"family":"Times New Roman","size":12,"color":"#000000"},"alignment":{"horizontal":"right","vertical":"center"},"border":[{"type":"left","color":"#000000","style":1},{"type":"top","color":"#000000","style":1},{"type":"right","color":"#000000","style":1},{"type":"bottom","color":"#000000","style":1}]}`
	subtotalNumberConditionStyle = `{"fill":{"type":"pattern","pattern":1,"color":["#E8E8E8"]},"alignment":{"horizontal":"right","vertical":"center"},"font":{"family":"Times New Roman","bold":false,"size":12,"color":"#FF0000"},"border":[{"type":"left","color":"#000000","style":1},{"type":"top","color":"#000000","style":1},{"type":"right","color":"#000000","style":1},{"type":"bottom","color":"#000000","style":1}]}`
	subtotalDecimalsStyle        = `{"fill":{"type":"pattern","pattern":1,"color":["#E8E8E8"]},"number_format":4,"font":{"bold":false,"italic":false,"family":"Times New Roman","size":12,"color":"#000000"},"alignment":{"horizontal":"right","vertical":"center"},"border":[{"type":"left","color":"#000000","style":1},{"type":"top","color":"#000000","style":1},{"type":"right","color":"#000000","style":1},{"type":"bottom","color":"#000000","style":1}]}`
	subtotalDecimalsCondition    = `{"fill":{"type":"pattern","pattern":1,"color":["#E8E8E8"]},"number_format":4,"alignment":{"horizontal":"right","vertical":"center"},"font":{"family":"Times New Roman","bold":false,"size":12,"color":"#FF0000"},"border":[{"type":"left","color":"#000000","style":1},{"type":"top","color":"#000000","style":1},{"type":"right","color":"#000000","style":1},{"type":"bottom","color":"#000000","style":1}]}`
)

func init() {
	// 用于标题 例如明细总表大标题
	style.Register(style.Title, style.NewStyle(titleStyle))
	// 用于头部
	style.Register(style.Head, style.NewStyle(headStyle))
	// 用于普通文本
	style.Register(style.Text, style.NewStyle(textStyle))
	// 用于数字
	style.Register(style.Number, style.NewStyle(numberStyle))
	// 当符合条件时，数字使用条件样式
	style.Register(style.NumberCondition, style.NewConditionStyle(numberConditionStyle))
	// 用于小数
	style.Register(style.Decimals, style.NewStyle(decimalsStyle))
	// 用于小计（组合单位小计）
	style.Register(style.SubtotalText, style.NewStyle(subtotalTextStyle))
	// 当符合条件时，小数使用条件样式
	style.Register(style.DecimalsCondition, style.NewConditionStyle(decimalsConditionStyle))
	// 用于小计数量
	style.Register(style.SubtotalNumber, style.NewStyle(subtotalNumberStyle))
	// 当符合条件时，小计金额使用条件样式
	style.Register(style.SubtotalNumberCondition, style.NewConditionStyle(subtotalNumberConditionStyle))
	// 用于小计金额
	style.Register(style.SubtotalDecimals, style.NewStyle(subtotalDecimalsStyle))
	// 当符合条件时，小计金额使用条件样式
	style.Register(style.SubtotalDecimalsCondition, style.NewConditionStyle(subtotalDecimalsCondition))
}
