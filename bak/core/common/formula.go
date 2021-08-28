package common

type FormulaExpress string

func (express FormulaExpress) Value() string {
	return string(express)
}

const (
	// 求和
	SUM     FormulaExpress = "=SUM(%s:%s)"
	// 平均值
	AVERAGE FormulaExpress = "=AVERAGE(%s:%s)"
	// 数量
	COUNT   FormulaExpress = "=COUNT(%s:%s)"
	// 最大值
	MAX     FormulaExpress = "=MAX(%s:%s)"
	// 最小值
	MIN     FormulaExpress = "=MIN(%s:%s)"
	// 绝对值
	ABS     FormulaExpress = "=ABS(%v)"
)
