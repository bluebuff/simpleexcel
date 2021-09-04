package formula

import "fmt"

type Formula string

func (formula Formula) Format(params ...interface{}) string {
	return fmt.Sprintf(string(formula), params...)
}

const (
	SUM     Formula = "=SUM(%s:%s)"
	AVERAGE Formula = "=AVERAGE(%s:%s)"
	COUNT   Formula = "=COUNT(%s:%s)"
	MAX     Formula = "=MAX(%s:%s)"
	MIN     Formula = "=MIN(%s:%s)"
	ABS     Formula = "=ABS(%v)"
)
