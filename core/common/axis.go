package common

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize/v2"
)

type Axis struct {
	Col int
	Row int
}

func (axis Axis) String() string {
	colName, _ := excelize.ColumnNumberToName(axis.Col)
	return fmt.Sprintf("%s%d", colName, axis.Row)
}
