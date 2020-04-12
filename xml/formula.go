package xml

type FormulaFormatAlias string

const (
	FormulaSumAlias FormulaFormatAlias = "SUM"
)

type FormulaExpress string

const (
	FormulaSumExpress = "=SUM(%s:%s)"
)

var FormatMap = map[FormulaFormatAlias]string{
	FormulaSumAlias: FormulaSumExpress,
}
