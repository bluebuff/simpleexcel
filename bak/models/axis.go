package models

import (
	"fmt"
	"github.com/xuri/excelize/v2"
)

type Axis struct {
	Col int
	Row int
}

func (axis Axis) String() string {
	colName, _ := excelize.ColumnNumberToName(axis.Col)
	return fmt.Sprintf("%s%d", colName, axis.Row)
}
