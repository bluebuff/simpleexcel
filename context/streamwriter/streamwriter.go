package streamwriter

import (
	"fmt"
	"github.com/bluebuff/simpleexcel/v2/context"
	"github.com/bluebuff/simpleexcel/v2/internal"
	"github.com/bluebuff/simpleexcel/v2/model"
	"github.com/bluebuff/simpleexcel/v2/option"
	"github.com/bluebuff/simpleexcel/v2/style"
	"github.com/xuri/excelize/v2"
)

func NewContext(streamWriter *excelize.StreamWriter, styleMng style.StyleManager) context.Context {
	layout := model.Layout{Left: 1, Top: 1, Right: 2, Bottom: 1}
	return &streamWriterContext{
		layout:       layout,
		streamWriter: streamWriter,
		rowCounter:   internal.NewCounter(layout.Top),
		colCounter:   internal.NewCounter(layout.Left),
		cells:        make(model.Cells, 0, 50),
		styleMng:     styleMng,
	}
}

type streamWriterContext struct {
	context.Context
	layout       model.Layout
	streamWriter *excelize.StreamWriter
	cells        model.Cells
	rowCounter   internal.Counter
	colCounter   internal.Counter
	styleMng     style.StyleManager
}

func (ctx *streamWriterContext) SetLayout(layout model.Layout) {
	if layout.Left != 0 {
		ctx.layout.Left = layout.Left
		ctx.colCounter = internal.NewCounter(layout.Left)
	}
	if layout.Right != 0 {
		ctx.layout.Right = layout.Right
	}
	if layout.Top != 0 {
		ctx.layout.Top = layout.Top
		ctx.rowCounter = internal.NewCounter(layout.Top)
	}
	if layout.Bottom != 0 {
		ctx.layout.Bottom = layout.Bottom
	}
}

func (ctx *streamWriterContext) GetLayout() model.Layout {
	return ctx.layout
}

func (ctx *streamWriterContext) SetColWidth(startIndex, endIndex int, width float64) {
	ctx.streamWriter.SetColWidth(startIndex, endIndex, width)
	if ctx.layout.Right < endIndex+1 {
		ctx.layout.Right = endIndex + 1
	}
}

func (ctx *streamWriterContext) SetTitle(value string, opts ...interface{}) string {
	styleId, _ := ctx.styleMng.Get(style.Title)
	cell := &model.Cell{Value: value, StyleID: styleId}
	return ctx.SetInterface(cell, opts...)
}

func (ctx *streamWriterContext) SetTitleLine(value string, opts ...interface{}) string {
	startAxis := ctx.CurrentAxis()
	for i := ctx.CurrentColCount(); i < ctx.layout.Right; i++ {
		ctx.SetTitle(value, opts...)
	}
	endAxis := ctx.GetHorizontalAxis(ctx.layout.Right - ctx.CurrentColCount() - 1)
	ctx.MergeValue(startAxis, endAxis)
	return endAxis
}

func (ctx *streamWriterContext) SetHeader(value string, opts ...interface{}) string {
	styleId, _ := ctx.styleMng.Get(style.Head)
	cell := &model.Cell{Value: value, StyleID: styleId}
	return ctx.SetInterface(cell, opts...)
}

func (ctx *streamWriterContext) SetHeaders(headers []string, opts ...interface{}) string {
	var last string
	for _, value := range headers {
		last = ctx.SetHeader(value, opts...)
	}
	return last
}

func (ctx *streamWriterContext) SetFormula(formula string, opts ...interface{}) string {
	styleId, _ := ctx.styleMng.Get(style.SubtotalText)
	cell := &model.Cell{StyleID: styleId, Formula: formula}
	return ctx.SetInterface(cell, opts...)
}

func (ctx *streamWriterContext) NewLine() {
	ctx.flush()
	ctx.colCounter.Reset()
	ctx.rowCounter.Incr()
	ctx.layout.Bottom = ctx.CurrentRowCount()
}

func (ctx *streamWriterContext) flush() {
	values := ctx.cells.ToInterface()
	if len(values) != 0 {
		axis := ctx.GetHorizontalAxis(-len(values))
		ctx.streamWriter.SetRow(axis, values)
		ctx.cells = ctx.cells[:0]
	}
}

func (ctx *streamWriterContext) MergeValue(hcell, vcell string) {
	ctx.streamWriter.MergeCell(hcell, vcell)
}

func (ctx *streamWriterContext) SetInt32(value int32, opts ...interface{}) string {
	styleId, _ := ctx.styleMng.Get(style.Number)
	cell := &model.Cell{Value: value, StyleID: styleId}
	return ctx.SetInterface(cell, opts...)
}

func (ctx *streamWriterContext) SetInt64(value int64, opts ...interface{}) string {
	styleId, _ := ctx.styleMng.Get(style.Number)
	cell := &model.Cell{Value: value, StyleID: styleId}
	return ctx.SetInterface(cell, opts...)
}

func (ctx *streamWriterContext) SetUint32(value uint32, opts ...interface{}) string {
	styleId, _ := ctx.styleMng.Get(style.Number)
	cell := &model.Cell{Value: value, StyleID: styleId}
	return ctx.SetInterface(cell, opts...)
}

func (ctx *streamWriterContext) SetUint64(value uint64, opts ...interface{}) string {
	styleId, _ := ctx.styleMng.Get(style.Number)
	cell := &model.Cell{Value: value, StyleID: styleId}
	return ctx.SetInterface(cell, opts...)
}

// 写入float32类型的数据
func (ctx *streamWriterContext) SetFloat32(value float32, opts ...interface{}) string {
	styleId, _ := ctx.styleMng.Get(style.Decimals)
	cell := &model.Cell{Value: value, StyleID: styleId}
	return ctx.SetInterface(cell, opts...)
}

// 写入float64类型的数据
func (ctx *streamWriterContext) SetFloat64(value float64, opts ...interface{}) string {
	styleId, _ := ctx.styleMng.Get(style.Decimals)
	cell := &model.Cell{Value: value, StyleID: styleId}
	return ctx.SetInterface(cell, opts...)
}

// 写入string类型的数据
func (ctx *streamWriterContext) SetString(value string, opts ...interface{}) string {
	styleId, _ := ctx.styleMng.Get(style.Text)
	cell := &model.Cell{Value: value, StyleID: styleId}
	return ctx.SetInterface(cell, opts...)
}

// 设置一行的合并的字符串
func (ctx *streamWriterContext) SetStringLine(value string, opts ...interface{}) string {
	startAxis := ctx.CurrentAxis()
	for i := ctx.CurrentColCount(); i < ctx.layout.Right; i++ {
		ctx.SetString(value, opts...)
	}
	endAxis := ctx.GetHorizontalAxis(ctx.layout.Right - ctx.CurrentColCount() - 1)
	ctx.MergeValue(startAxis, endAxis)
	return endAxis
}

// 写入interface的数据，无样式
func (ctx *streamWriterContext) SetInterface(cell *model.Cell, opts ...interface{}) string {
	// do options
	ctx.doOptions(cell, opts...)
	ctx.cells = append(ctx.cells, cell)
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

func (ctx *streamWriterContext) doOptions(cell *model.Cell, opts ...interface{}) {
	for _, opt := range opts {
		if opt == nil {
			continue
		}
		switch v := opt.(type) {
		case option.Option:
			v(ctx.styleMng, cell)
		}
	}
	return
}

func (ctx *streamWriterContext) String() string {
	format := `+-------------%02d--------------+
|                             |
%02d                           %02d
|                             |
+-------------%02d--------------+`
	return fmt.Sprintf(format, ctx.layout.Top, ctx.layout.Left, ctx.layout.Right, ctx.layout.Bottom)
}
