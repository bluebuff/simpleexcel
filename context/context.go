package context

import (
	"github.com/bluebuff/simpleexcel/v2/model"
)

type Context interface {
	// 设置列的宽度
	SetColWidth(startIndex, endIndex int, width float64)

	// 设置边界的位置，上下左右的偏移量，从1开始
	SetLayout(layout model.Layout)

	GetLayout() model.Layout

	// 设置大标题
	SetTitle(value string, opts ...interface{}) string

	// 设置一行的大标题
	SetTitleLine(value string, opts ...interface{}) string

	// 设置表头
	SetHeader(value string, opts ...interface{}) string

	// 设置表头列表
	SetHeaders(headers []string, opts ...interface{}) string

	// 设置公式
	SetFormula(formula string, opts ...interface{}) string

	// 换行
	NewLine()

	// 合并单元格
	MergeValue(hcell, vcell string)

	// 写入int32类型的数据
	SetInt32(value int32, opts ...interface{}) string

	// 写入int64类型的数据
	SetInt64(value int64, opts ...interface{}) string

	// 写入uint32类型的数据
	SetUint32(value uint32, opts ...interface{}) string

	// 写入uint64类型的数据
	SetUint64(value uint64, opts ...interface{}) string

	// 写入float32类型的数据
	SetFloat32(value float32, opts ...interface{}) string

	// 写入float64类型的数据
	SetFloat64(value float64, opts ...interface{}) string

	// 写入string类型的数据
	SetString(value string, opts ...interface{}) string

	// 设置一行的合并的字符串
	SetStringLine(value string, opts ...interface{}) string

	// 写入interface的数据，无样式
	SetInterface(cell *model.Cell, opts ...interface{}) string

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

	String() string
}
