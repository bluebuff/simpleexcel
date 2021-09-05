package simpleexcel

import (
	"fmt"
	"github.com/bluebuff/simpleexcel/v2/context"
	"github.com/bluebuff/simpleexcel/v2/option"
	"github.com/bluebuff/simpleexcel/v2/style"
	"github.com/bluebuff/simpleexcel/v2/style/standard"
	"math/rand"
	"testing"
)

func TestNewStreamWriterExcelBuilder(t *testing.T) {
	builder := NewStreamWriterExcelBuilder(standard.Style)
	builder.JoinSheet("学生成绩表", do)
	builder.JoinSheet("批量数据", do2)
	builder.Active("批量数据")
	fileName, err := builder.Build()
	if err != nil {
		fmt.Println(err)
		return
	}
	fmt.Println(fileName)
}

func do2(ctx context.Context) error {

	ctx.SetColWidth(1, 60, 20)

	for i := 0; i < 50; i++ {
		ctx.SetHeader(fmt.Sprintf("列%d", i+1))
	}

	ctx.NewLine()

	for row := 2; row <= 500; row++ {
		for i := 0; i < 60; i++ {
			ctx.SetString(fmt.Sprint(rand.Int()))
		}
		ctx.NewLine()
		if (row-1)%5 == 0 {
			ctx.MergeValue(fmt.Sprintf("A%d", row-4), fmt.Sprintf("A%d", row))
		}
	}

	return nil
}

func do(ctx context.Context) error {
	ctx.SetColWidth(1, 3, 20)
	oo := ctx.SetTitleLine("2021年学分表")
	fmt.Println(oo)
	ctx.NewLine()
	a := ctx.SetHeader("姓名")
	fmt.Println(a)
	b := ctx.SetHeader("年龄")
	fmt.Println(b)
	c := ctx.SetHeader("学分")
	fmt.Println(c)
	ctx.NewLine()
	e := ctx.SetString("张三")
	fmt.Println(e)
	ctx.SetUint32(26)
	ctx.SetFloat32(98)
	firstAxis := ctx.LastAxis()
	ctx.NewLine()
	ctx.SetString("李四")
	ctx.SetUint32(55, option.CompareAndStyleUint32(option.LE, 55))
	ctx.SetFloat32(100, option.CompareAndStyleFloat32(option.EQ, 100))
	lastAxis := ctx.LastAxis()
	ctx.NewLine()
	ctx.SetStringLine("")
	ctx.NewLine()
	d := ctx.SetHeader("小计")
	fmt.Println(d)
	ctx.SetHeader("")
	f := ctx.SetFormula(fmt.Sprintf("=SUM(%s:%s)", firstAxis, lastAxis), option.CustomStyleByAlias(style.SubtotalDecimalsCondition))
	fmt.Println(f)
	ctx.NewLine()
	g := ctx.SetStringLine("test")
	fmt.Println(g)
	ctx.NewLine()
	hcell := ctx.SetString("", option.WhenEmptyAndSwapString("-"))
	ctx.SetString("Cell1-")
	ctx.SetString("Cell1--")
	ctx.NewLine()
	vcell := ctx.SetString("Cell2")
	ctx.SetString("Cell2-")
	ctx.SetString("Cell2--")
	ctx.NewLine()
	ctx.MergeValue(hcell, vcell)
	return nil
}
