package simpleexcel

import (
	"github.com/xuri/excelize/v2"
)

type Cell = excelize.Cell

type Cells []*Cell

func (c Cells) ToInterface() []interface{} {
	s := make([]interface{}, 0, len(c))
	for _, item := range c {
		s = append(s, item)
	}
	return s
}
