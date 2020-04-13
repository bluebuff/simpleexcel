package context

import (
	"errors"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/bluebuff/simple-excelize/core/common"
	"github.com/bluebuff/simple-excelize/models"
)

type Handler func(ctx Context) error

type Context interface {
	// 设置列的宽度
	SetColWidth(startIndex, endIndex int, width float64)

	// 设置边界的位置，上下左右的偏移量，从1开始
	SetLayout(layout *models.Layout)

	// 设置大标题
	SetTitle(value string, useStyle ...bool) string

	// 设置一行的大标题
	SetTitleLine(value string, useStyle ...bool) string

	// 设置表头
	SetHeader(value string, useStyle ...bool) string

	// 设置表头列表
	SetHeaders(headers []string, useStyle ...bool) string

	// 设置公式
	SetFormula(formula string, handle Handler) string

	// 换行
	NewLine()

	// 合并单元格
	MergeValue(hcell, vcell string)

	// 写入int32类型的数据
	SetInt32(value int32, useStyle ...bool) string

	// 写入int64类型的数据
	SetInt64(value int64, useStyle ...bool) string

	// 写入uint32类型的数据
	SetUint32(value uint32, useStyle ...bool) string

	// 写入uint64类型的数据
	SetUint64(value uint64, useStyle ...bool) string

	// 写入float32类型的数据
	SetFloat32(value float32, useStyle ...bool) string

	// 写入float64类型的数据
	SetFloat64(value float64, useStyle ...bool) string

	// 写入string类型的数据
	SetString(value string, useStyle ...bool) string

	// 设置一行的合并的字符串
	SetStringLine(value string, useStyle ...bool) string

	// 写入interface的数据，无样式
	SetInterface(value interface{}) string

	// 横向计数器向后偏移一段距离，默认为1
	Jump(offset ...int) string

	// 获取最近的单元格坐标名称，例如：A1、B1
	LastAxis() string

	// 获取当前位置的偏移后的坐标
	GetAxis(hoffset, voffset int) string

	// 获取当前位置的偏移后（横向）的坐标
	GetHorizontalAxis(offset int) string

	//  获取当前位置的偏移后（纵向）的坐标
	GetVerticalAxis(offset int) string

	// 获取当前位置的坐标
	CurrentAxis() string

	// 获取当前行数
	CurrentRowCount() int

	// 获取当前列数
	CurrentColCount() int

	// 获取指定单元坐标值
	GetValue(axis string) string

	// 设置样式
	SetStyle(styleName common.StyleNameAlias, useStyle ...bool)

	// 设置条件样式
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

func (ctx *context) SetInt32(value int32, useStyle ...bool) string {
	ctx.SetStyle(common.Number, useStyle...)
	ctx.SetConditionStyle(common.NumberCondition, common.NumberConditionExpression, 0, useStyle...)
	return ctx.SetInterface(value)
}

func (ctx *context) SetInt64(value int64, useStyle ...bool) string {
	ctx.SetStyle(common.Number, useStyle...)
	ctx.SetConditionStyle(common.NumberCondition, common.NumberConditionExpression, 0, useStyle...)
	return ctx.SetInterface(value)
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
