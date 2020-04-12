package context

import (
	"errors"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/bluebuff/simple-excelize/v1/core/common"
	"github.com/bluebuff/simple-excelize/v1/models"
)

type Handler func(ctx Context) error

type Context interface {
	SetColWidth(startIndex, endIndex int, width float64)

	SetLayout(layout *models.Layout)

	SetTitle(value string, useStyle ...bool) string

	SetTitleLine(value string, useStyle ...bool) string

	SetHeader(value string, useStyle ...bool) string

	SetHeaders(headers []string, useStyle ...bool) string

	SetFormula(formula string, handle Handler) string

	NewLine()

	MergeValue(hcell, vcell string)

	SetUint32(value uint32, useStyle ...bool) string

	SetUint64(value uint64, useStyle ...bool) string

	SetFloat32(value float32, useStyle ...bool) string

	SetFloat64(value float64, useStyle ...bool) string

	SetString(value string, useStyle ...bool) string

	SetStringLine(value string, useStyle ...bool) string

	SetInterface(value interface{}) string

	Jump(offset ...int) string

	LastAxis() string

	GetAxis(hoffset, voffset int) string

	GetHorizontalAxis(offset int) string

	GetVerticalAxis(offset int) string

	CurrentAxis() string

	CurrentRowCount() int

	CurrentColCount() int

	GetValue(axis string) string

	SetStyle(styleName common.StyleNameAlias, useStyle ...bool)

	SetConditionStyle(styleName common.StyleNameAlias, conditionExpression common.ConditionExpress, compareValue float64, useStyle ...bool)

	String() string
}

func NewContext(file *excelize.File, styleManager *common.StyleManager, sheetName string, layout models.Layout) Context {
	ctx := &context{
		file:         file,
		sheetName:    sheetName,
		styleManager: styleManager,
		rowCounter:   common.NewCounter(layout.Top),
		colCounter:   common.NewCounter(layout.Left),
		Layout:       layout,
	}
	return ctx
}

type context struct {
	file         *excelize.File
	styleManager *common.StyleManager
	sheetName    string
	rowCounter   common.Counter
	colCounter   common.Counter
	Layout       models.Layout
	Context
}

func (ctx *context) SetLayout(layout *models.Layout) {
	if layout == nil {
		return
	}
	if layout.Left != 0 {
		ctx.Layout.Left = layout.Left
		ctx.colCounter = common.NewCounter(layout.Left)
	}
	if layout.Right != 0 {
		ctx.Layout.Right = layout.Right
	}
	if layout.Top != 0 {
		ctx.Layout.Top = layout.Top
		ctx.rowCounter = common.NewCounter(layout.Top)
	}
	if layout.Bottom != 0 {
		ctx.Layout.Bottom = layout.Bottom
	}
}

func (ctx *context) SetColWidth(startIndex, endIndex int, width float64) {
	startCol, _ := excelize.ColumnNumberToName(startIndex)
	endCol, _ := excelize.ColumnNumberToName(endIndex)
	ctx.file.SetColWidth(ctx.sheetName, startCol, endCol, width)
	if ctx.Layout.Right < endIndex+1 {
		ctx.Layout.Right = endIndex + 1
	}
}

func (ctx *context) SetTitle(value string, useStyle ...bool) string {
	ctx.SetStyle(common.Title, useStyle...)
	return ctx.SetInterface(value)
}

func (ctx *context) SetTitleLine(value string, useStyle ...bool) string {
	startAxis := ctx.SetTitle(value, useStyle...)
	endAxis := ctx.GetHorizontalAxis(ctx.Layout.Right - ctx.CurrentColCount() - 1)
	ctx.MergeValue(startAxis, endAxis)
	return endAxis
}

func (ctx *context) SetHeader(value string, useStyle ...bool) string {
	ctx.SetStyle(common.Head, useStyle...)
	return ctx.SetInterface(value)
}

func (ctx *context) SetHeaders(headers []string, useStyle ...bool) string {
	var last string
	for _, value := range headers {
		last = ctx.SetHeader(value, useStyle...)
	}
	return last
}

func (ctx *context) SetFormula(formula string, handle Handler) string {
	handle(ctx)
	ctx.file.SetCellFormula(ctx.sheetName, ctx.CurrentAxis(), formula)
	return ctx.Jump()
}

func (ctx *context) NewLine() {
	ctx.colCounter.Reset()
	ctx.rowCounter.Incr()
	ctx.Layout.Bottom = ctx.CurrentRowCount()
}

func (ctx *context) MergeValue(hcell, vcell string) {
	index, _ := ctx.file.GetCellStyle(ctx.sheetName, hcell)
	ctx.file.MergeCell(ctx.sheetName, hcell, vcell)
	ctx.file.SetCellStyle(ctx.sheetName, hcell, vcell, index)
	right, _ := excelize.ColumnNameToNumber(vcell)
	if ctx.Layout.Right < right {
		ctx.Layout.Right = right
	}
}

func (ctx *context) SetUint32(value uint32, useStyle ...bool) string {
	ctx.SetStyle(common.Number, useStyle...)
	ctx.SetConditionStyle(common.NumberCondition, common.NumberConditionExpression, 0, useStyle...)
	return ctx.SetInterface(value)
}

func (ctx *context) SetUint64(value uint64, useStyle ...bool) string {
	ctx.SetStyle(common.Number, useStyle...)
	ctx.SetConditionStyle(common.NumberCondition, common.NumberConditionExpression, 0, useStyle...)
	return ctx.SetInterface(value)
}

func (ctx *context) SetFloat32(value float32, useStyle ...bool) string {
	ctx.SetStyle(common.Decimals, useStyle...)
	ctx.SetConditionStyle(common.DecimalsCondition, common.DecimalsConditionExpression, 0, useStyle...)
	return ctx.SetInterface(value)
}

func (ctx *context) SetFloat64(value float64, useStyle ...bool) string {
	ctx.SetStyle(common.Decimals, useStyle...)
	ctx.SetConditionStyle(common.DecimalsCondition, common.DecimalsConditionExpression, 0, useStyle...)
	return ctx.SetInterface(value)
}

func (ctx *context) SetString(value string, useStyle ...bool) string {
	ctx.SetStyle(common.Text, useStyle...)
	return ctx.SetInterface(value)
}

func (ctx *context) SetStringLine(value string, useStyle ...bool) string {
	startAxis := ctx.SetString(value, useStyle...)
	endAxis := ctx.GetHorizontalAxis(ctx.Layout.Right - ctx.CurrentColCount() - 1)
	ctx.MergeValue(startAxis, endAxis)
	return endAxis
}

func (ctx *context) GetValue(axis string) string {
	value, err := ctx.file.GetCellValue(ctx.sheetName, axis)
	if err == nil {
		return value
	}
	return ""
}

func (ctx *context) SetInterface(value interface{}) string {
	ctx.file.SetCellValue(ctx.sheetName, ctx.CurrentAxis(), value)
	return ctx.Jump()
}

func (ctx *context) Jump(offset ...int) string {
	var before int
	if offset == nil || len(offset) == 0 {
		before = ctx.colCounter.Incr()
	} else {
		before = ctx.colCounter.IncrBy(offset[0])
	}
	right := ctx.CurrentColCount()
	if ctx.Layout.Right < right {
		ctx.Layout.Right = right
	}
	col, _ := excelize.ColumnNumberToName(before)
	return fmt.Sprintf("%s%d", col, ctx.CurrentRowCount())
}

func (ctx *context) LastAxis() string {
	axis, _ := excelize.ColumnNumberToName(ctx.CurrentColCount() - 1)
	return fmt.Sprintf("%s%d", axis, ctx.CurrentRowCount())
}

func (ctx *context) GetAxis(hoffset, voffset int) string {
	axis, _ := excelize.ColumnNumberToName(ctx.CurrentColCount() + hoffset)
	return fmt.Sprintf("%s%d", axis, ctx.CurrentRowCount()+voffset)
}

func (ctx *context) CurrentAxis() string {
	return ctx.GetAxis(0, 0)
}

func (ctx *context) GetHorizontalAxis(offset int) string {
	return ctx.GetAxis(offset, 0)
}

func (ctx *context) GetVerticalAxis(offset int) string {
	return ctx.GetAxis(0, offset)
}

func (ctx *context) CurrentRowCount() int {
	return ctx.rowCounter.Current()
}

func (ctx *context) CurrentColCount() int {
	return ctx.colCounter.Current()
}

func (ctx *context) SetStyle(styleName common.StyleNameAlias, useStyle ...bool) {
	if styleFunc, ok := ctx.styleManager.GetStyleFunc(styleName); ok && (useStyle == nil || len(useStyle) == 0 || useStyle[0]) {
		styleFunc(ctx.sheetName, ctx.CurrentAxis(), ctx.CurrentAxis(), string(common.NotConditionExpression), 0)
	}
}

func (ctx *context) SetConditionStyle(styleName common.StyleNameAlias, conditionExpression common.ConditionExpress, compareValue float64, useStyle ...bool) {
	if styleFunc, ok := ctx.styleManager.GetStyleFunc(styleName); ok && (useStyle == nil || len(useStyle) == 0 || useStyle[0]) {
		styleFunc(ctx.sheetName, ctx.CurrentAxis(), ctx.CurrentAxis(), string(conditionExpression), compareValue)
	}
}

func SetConditionStyle(alias common.StyleNameAlias, compareValue float64) Handler {
	return func(ctx Context) error {
		ctx.SetStyle(alias)
		switch alias {
		case common.SubtotalNumber:
			ctx.SetConditionStyle(common.SubtotalNumberCondition, common.NumberConditionExpression, compareValue)
		case common.SubtotalDecimals:
			ctx.SetConditionStyle(common.SubtotalDecimalsCondition, common.DecimalsConditionExpression, compareValue)
		default:
			return errors.New("alias not support")
		}
		return nil
	}
}

func (ctx *context) String() string {
	format := `+-------------%02d--------------+
|                             |
%02d                           %02d
|                             |
+-------------%02d--------------+`
	return fmt.Sprintf(format, ctx.Layout.Top, ctx.Layout.Left, ctx.Layout.Right, ctx.Layout.Bottom)
}
