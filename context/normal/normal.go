package normal

import (
	"fmt"
	"github.com/bluebuff/simpleexcel/v2/context"
	"github.com/bluebuff/simpleexcel/v2/internal"
	"github.com/bluebuff/simpleexcel/v2/model"
	"github.com/bluebuff/simpleexcel/v2/option"
	"github.com/bluebuff/simpleexcel/v2/style"
	"github.com/xuri/excelize/v2"
)

func NewContext(file *excelize.File, styleMng style.StyleManager, sheetName string) context.Context {
	layout := model.Layout{Left: 1, Top: 1, Right: 2, Bottom: 1}
	ctx := &normalContext{
		file:       file,
		sheetName:  sheetName,
		styleMng:   styleMng,
		layout:     layout,
		rowCounter: internal.NewCounter(layout.Top),
		colCounter: internal.NewCounter(layout.Left),
	}
	return ctx
}

type normalContext struct {
	file       *excelize.File
	sheetName  string
	styleMng   style.StyleManager
	rowCounter internal.Counter
	colCounter internal.Counter
	layout     model.Layout
}

func (ctx *normalContext) SetLayout(layout model.Layout) {
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

func (ctx *normalContext) GetLayout() model.Layout {
	return ctx.layout
}

func (ctx *normalContext) SetColWidth(startIndex, endIndex int, width float64) {
	startCol, _ := excelize.ColumnNumberToName(startIndex)
	endCol, _ := excelize.ColumnNumberToName(endIndex)
	ctx.file.SetColWidth(ctx.sheetName, startCol, endCol, width)
	if ctx.layout.Right < endIndex+1 {
		ctx.layout.Right = endIndex + 1
	}
}

func (ctx *normalContext) SetTitle(value string, opts ...interface{}) string {
	styleId, _ := ctx.styleMng.Get(style.Title)
	cell := &model.Cell{Value: value, StyleID: styleId}
	return ctx.SetInterface(cell, opts...)
}

func (ctx *normalContext) SetTitleLine(value string, opts ...interface{}) string {
	startAxis := ctx.SetTitle(value, opts...)
	endAxis := ctx.GetHorizontalAxis(ctx.layout.Right - ctx.CurrentColCount() - 1)
	ctx.MergeValue(startAxis, endAxis)
	return endAxis
}

func (ctx *normalContext) SetHeader(value string, opts ...interface{}) string {
	styleId, _ := ctx.styleMng.Get(style.Head)
	cell := &model.Cell{Value: value, StyleID: styleId}
	return ctx.SetInterface(cell, opts...)
}

func (ctx *normalContext) SetHeaders(headers []string, opts ...interface{}) string {
	var last string
	for _, value := range headers {
		last = ctx.SetHeader(value, opts...)
	}
	return last
}

func (ctx *normalContext) SetFormula(formula string, opts ...interface{}) string {
	styleId, _ := ctx.styleMng.Get(style.SubtotalText)
	cell := &model.Cell{StyleID: styleId, Formula: formula}
	return ctx.SetInterface(cell, opts...)
}

func (ctx *normalContext) NewLine() {
	ctx.colCounter.Reset()
	ctx.rowCounter.Incr()
	ctx.layout.Bottom = ctx.CurrentRowCount()
}

func (ctx *normalContext) MergeValue(hcell, vcell string) {
	index, _ := ctx.file.GetCellStyle(ctx.sheetName, hcell)
	ctx.file.MergeCell(ctx.sheetName, hcell, vcell)
	ctx.file.SetCellStyle(ctx.sheetName, hcell, vcell, index)
	right, _ := excelize.ColumnNameToNumber(vcell)
	if ctx.layout.Right < right {
		ctx.layout.Right = right
	}
}

func (ctx *normalContext) SetInt32(value int32, opts ...interface{}) string {
	styleId, _ := ctx.styleMng.Get(style.Number)
	cell := &model.Cell{Value: value, StyleID: styleId}
	return ctx.SetInterface(cell, opts...)
}

func (ctx *normalContext) SetInt64(value int64, opts ...interface{}) string {
	styleId, _ := ctx.styleMng.Get(style.Number)
	cell := &model.Cell{Value: value, StyleID: styleId}
	return ctx.SetInterface(cell, opts...)
}

func (ctx *normalContext) SetUint32(value uint32, opts ...interface{}) string {
	styleId, _ := ctx.styleMng.Get(style.Number)
	cell := &model.Cell{Value: value, StyleID: styleId}
	return ctx.SetInterface(cell, opts...)
}

func (ctx *normalContext) SetUint64(value uint64, opts ...interface{}) string {
	styleId, _ := ctx.styleMng.Get(style.Number)
	cell := &model.Cell{Value: value, StyleID: styleId}
	return ctx.SetInterface(cell, opts...)
}

func (ctx *normalContext) SetFloat32(value float32, opts ...interface{}) string {
	styleId, _ := ctx.styleMng.Get(style.Decimals)
	cell := &model.Cell{Value: value, StyleID: styleId}
	return ctx.SetInterface(cell, opts...)
}

func (ctx *normalContext) SetFloat64(value float64, opts ...interface{}) string {
	styleId, _ := ctx.styleMng.Get(style.Decimals)
	cell := &model.Cell{Value: value, StyleID: styleId}
	return ctx.SetInterface(cell, opts...)
}

func (ctx *normalContext) SetString(value string, opts ...interface{}) string {
	styleId, _ := ctx.styleMng.Get(style.Text)
	cell := &model.Cell{Value: value, StyleID: styleId}
	return ctx.SetInterface(cell, opts...)
}

func (ctx *normalContext) SetStringLine(value string, opts ...interface{}) string {
	startAxis := ctx.SetString(value, opts...)
	endAxis := ctx.GetHorizontalAxis(ctx.layout.Right - ctx.CurrentColCount() - 1)
	ctx.MergeValue(startAxis, endAxis)
	return endAxis
}

func (ctx *normalContext) GetValue(axis string) string {
	value, err := ctx.file.GetCellValue(ctx.sheetName, axis)
	if err == nil {
		return value
	}
	return ""
}

func (ctx *normalContext) SetInterface(cell *model.Cell, opts ...interface{}) string {
	// do options
	ctx.doOptions(cell, opts...)
	ctx.file.SetCellStyle(ctx.sheetName, ctx.CurrentAxis(), ctx.CurrentAxis(), cell.StyleID)
	if cell.Value != nil {
		ctx.file.SetCellValue(ctx.sheetName, ctx.CurrentAxis(), cell.Value)
	}
	if cell.Formula != "" {
		ctx.file.SetCellFormula(ctx.sheetName, ctx.CurrentAxis(), cell.Formula)
	}
	return ctx.Jump()
}

func (ctx *normalContext) Jump(offset ...int) string {
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

func (ctx *normalContext) LastAxis() string {
	axis, _ := excelize.ColumnNumberToName(ctx.CurrentColCount() - 1)
	return fmt.Sprintf("%s%d", axis, ctx.CurrentRowCount())
}

func (ctx *normalContext) GetAxis(hoffset, voffset int) string {
	axis, _ := excelize.ColumnNumberToName(ctx.CurrentColCount() + hoffset)
	return fmt.Sprintf("%s%d", axis, ctx.CurrentRowCount()+voffset)
}

func (ctx *normalContext) CurrentAxis() string {
	return ctx.GetAxis(0, 0)
}

func (ctx *normalContext) GetHorizontalAxis(offset int) string {
	return ctx.GetAxis(offset, 0)
}

func (ctx *normalContext) GetVerticalAxis(offset int) string {
	return ctx.GetAxis(0, offset)
}

func (ctx *normalContext) CurrentRowCount() int {
	return ctx.rowCounter.Current()
}

func (ctx *normalContext) CurrentColCount() int {
	return ctx.colCounter.Current()
}

func (ctx *normalContext) doOptions(cell *model.Cell, opts ...interface{}) {
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

func (ctx *normalContext) String() string {
	format := `+-------------%02d--------------+
|                             |
%02d                           %02d
|                             |
+-------------%02d--------------+`
	return fmt.Sprintf(format, ctx.layout.Top, ctx.layout.Left, ctx.layout.Right, ctx.layout.Bottom)
}
