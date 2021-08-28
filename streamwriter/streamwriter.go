package streamwriter

import (
	"fmt"
	"github.com/bluebuff/simpleexcel/v2"
	"github.com/bluebuff/simpleexcel/v2/internal"
	"github.com/bluebuff/simpleexcel/v2/style"
	"github.com/xuri/excelize/v2"
)

func New(streamWriter *excelize.StreamWriter) simpleexcel.Context {
	return &streamWriterContext{
		streamWriter: streamWriter,
		cells:        make(Cells, 0),
	}
}

type Cell excelize.Cell

type Cells []Cell

func (c Cells) ToInterface() []interface{} {
	s := make([]interface{}, 0, len(c))
	for _, item := range c {
		s = append(s, item)
	}
	return s
}

type streamWriterContext struct {
	simpleexcel.Context
	layout       simpleexcel.Layout
	streamWriter *excelize.StreamWriter
	cells        Cells
	rowCounter   internal.Counter
	colCounter   internal.Counter
	styleMng     *style.StyleManager
}

func (*streamWriterContext) SetTitle(value string, opts ...interface{}) string {

}

func (*streamWriterContext) SetTitleLine(value string, opts ...interface{}) string {

}

func (*streamWriterContext) SetHeader(value string, opts ...interface{}) string {

}

func (*streamWriterContext) SetHeaders(headers []string, opts ...interface{}) string {

}

func (*streamWriterContext) SetFormula(formula string) string {

}

func (ctx *streamWriterContext) NewLine() {
	axis := ctx.CurrentAxis()
	ctx.streamWriter.SetRow(axis, ctx.cells.ToInterface())
	ctx.cells = ctx.cells[:0]
}

func (ctx *streamWriterContext) MergeValue(hcell, vcell string) {
	ctx.streamWriter.MergeCell(hcell, vcell)
}

func (ctx *streamWriterContext) SetInt32(value int32, opts ...interface{}) string {
	var styleId int
	if fc, ok := style.GetStyleFunc(style.Number); ok {
		styleId, _ = fc(ctx.streamWriter.File)
	}
	return ctx.SetInterface(value, styleId, "")
}

func (ctx *streamWriterContext) SetInt64(value int64) string {
	var styleId int
	if fc, ok := style.GetStyleFunc(style.Number); ok {
		styleId, _ = fc(ctx.streamWriter.File)
	}
	return ctx.SetInterface(value, styleId, "")
}

func (ctx *streamWriterContext) SetUint32(value uint32) string {
	var styleId int
	if fc, ok := style.GetStyleFunc(style.Number); ok {
		styleId, _ = fc(ctx.streamWriter.File)
	}
	return ctx.SetInterface(value, styleId, "")
}

func (ctx *streamWriterContext) SetUint64(value uint64) string {
	var styleId int
	if fc, ok := style.GetStyleFunc(style.Number); ok {
		styleId, _ = fc(ctx.streamWriter.File)
	}
	return ctx.SetInterface(value, styleId, "")
}

// 写入float32类型的数据
func (ctx *streamWriterContext) SetFloat32(value float32) string {
	var styleId int
	if fc, ok := style.GetStyleFunc(style.Decimals); ok {
		styleId, _ = fc(ctx.streamWriter.File)
	}
	return ctx.SetInterface(value, styleId, "")
}

// 写入float64类型的数据
func (ctx *streamWriterContext) SetFloat64(value float64) string {
	var styleId int
	if fc, ok := style.GetStyleFunc(style.Decimals); ok {
		styleId, _ = fc(ctx.streamWriter.File)
	}
	return ctx.SetInterface(value, styleId, "")
}

// 写入string类型的数据
func (ctx *streamWriterContext) SetString(value string) string {
	var styleId int
	if fc, ok := style.GetStyleFunc(style.Text); ok {
		styleId, _ = fc(ctx.streamWriter.File)
	}
	return ctx.SetInterface(value, styleId, "")
}

// 设置一行的合并的字符串
func (ctx *streamWriterContext) SetStringLine(value string, opts ...interface{}) string {
	startAxis := ctx.SetString(value, opts...)
	endAxis := ctx.GetHorizontalAxis(ctx.layout.Right - ctx.CurrentColCount() - 1)
	ctx.MergeValue(startAxis, endAxis)
	return endAxis
}

// 写入interface的数据，无样式
func (ctx *streamWriterContext) SetInterface(value interface{}, styleId int, formula string) string {
	ctx.cells = append(ctx.cells, Cell{
		StyleID: styleId,
		Formula: formula,
		Value:   value,
	})
	return ctx.Jump()
}

func (ctx *streamWriterContext) Jump(offset ...int) string {
	var before int
	if offset == nil || len(offset) == 0 {
		before = ctx.colCounter.Incr()
	} else {
		before = ctx.colCounter.IncrBy(offset[0])
	}
	right := ctx.CurrentColCount()
	if ctx.layout.Right < right {
		ctx.layout.Right = right
	}
	col, _ := excelize.ColumnNumberToName(before)
	return fmt.Sprintf("%s%d", col, ctx.CurrentRowCount())
}

func (ctx *streamWriterContext) LastAxis() string {
	axis, _ := excelize.ColumnNumberToName(ctx.CurrentColCount() - 1)
	return fmt.Sprintf("%s%d", axis, ctx.CurrentRowCount())
}

func (ctx *streamWriterContext) GetAxis(hoffset, voffset int) string {
	axis, _ := excelize.ColumnNumberToName(ctx.CurrentColCount() + hoffset)
	return fmt.Sprintf("%s%d", axis, ctx.CurrentRowCount()+voffset)
}

func (ctx *streamWriterContext) CurrentAxis() string {
	return ctx.GetAxis(0, 0)
}

func (ctx *streamWriterContext) GetHorizontalAxis(offset int) string {
	return ctx.GetAxis(offset, 0)
}

func (ctx *streamWriterContext) GetVerticalAxis(offset int) string {
	return ctx.GetAxis(0, offset)
}

func (ctx *streamWriterContext) CurrentRowCount() int {
	return ctx.rowCounter.Current()
}

func (ctx *streamWriterContext) CurrentColCount() int {
	return ctx.colCounter.Current()
}

func (ctx *streamWriterContext) parseOptions(opts ...interface{}) (styleFunc style.CallSetStyleFunc, formula string) {
	for _, opt := range opts {
		switch v := opt.(type) {
		case style.CallSetStyleFunc:
			styleFunc = v
		case string:
			formula = v
		}
	}
	return
}

func (ctx *streamWriterContext) SetStyle(alias style.StyleNameAlias) {
	ctx.cells[len(ctx.cells)-1].StyleID, _ = ctx.styleMng.Get(alias)
}

func (ctx *streamWriterContext) compareLessThanOnStyle(value interface{}, compareValue interface{}) (styleId int) {
	switch v := value.(type) {
	case int32:
		if v < compareValue.(int32) {
			styleId, _ = ctx.styleMng.Get(style.NumberCondition)
		}
	case int64:
		if v < compareValue.(int64) {
			styleId, _ = ctx.styleMng.Get(style.NumberCondition)
		}
	case uint32:
		if v < compareValue.(uint32) {
			styleId, _ = ctx.styleMng.Get(style.NumberCondition)
		}
	case uint64:
		if v < compareValue.(uint64) {
			styleId, _ = ctx.styleMng.Get(style.NumberCondition)
		}
	case float32:
		if v < compareValue.(float32) {
			styleId, _ = ctx.styleMng.Get(style.DecimalsCondition)
		}
	case float64:
		if v < compareValue.(float64) {
			styleId, _ = ctx.styleMng.Get(style.DecimalsCondition)
		}
	}
	return
}
