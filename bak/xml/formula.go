package xml

import "github.com/bluebuff/simple-excelize/core/common"

type FormulaFormatAlias string

const (
	FormulaSumAlias FormulaFormatAlias = "SUM"
)

var FormatMap = map[FormulaFormatAlias]string{
	FormulaSumAlias: common.SUM.Value(),
}
