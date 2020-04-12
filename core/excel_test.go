package core

import (
	"fmt"
	"github.com/bluebuff/simple-excelize/v1/core/common"
	"github.com/bluebuff/simple-excelize/v1/core/context"
	"github.com/bluebuff/simple-excelize/v1/models"
	"io/ioutil"
	"testing"
)

func TestExcelBuilder_Build(t *testing.T) {
	builder := NewExcelBuilder()
	builder.RegisterStyle()
	builder.JoinSheet("列表1", buildSheetFunc)
	builder.JoinSheet("列表2", buildSheetFunc)
	builder.Active("列表1")
	bytes, err := builder.Build()
	if err != nil {
		t.Fatal(err)
	}
	err = ioutil.WriteFile("./test.xlsx", bytes, 0666)
	if err != nil {
		t.Fatal(err)
	}
}

func buildSheetFunc(ctx context.Context) error {
	ctx.SetLayout(&models.Layout{
		Left:  2,
		Top:   2,
		Right: 5,
	})
	ctx.SetColWidth(1, 2, 50)
	ctx.SetTitleLine("学生信息表")
	ctx.NewLine()
	ctx.SetHeader("姓名")
	ctx.SetHeader("年龄")
	ctx.SetHeader("学分")
	ctx.NewLine()
	ctx.SetString("张三")
	ctx.SetUint32(26)
	ctx.SetFloat32(98)
	firstAxis := ctx.LastAxis()
	ctx.NewLine()
	ctx.SetString("李四")
	ctx.SetUint32(28)
	ctx.SetFloat32(100)
	lastAxis := ctx.LastAxis()
	ctx.NewLine()
	ctx.SetStringLine("")
	ctx.NewLine()
	ctx.SetHeader("小计")
	ctx.SetHeader("")
	ctx.SetFormula(fmt.Sprintf("=SUM(%s:%s)", firstAxis, lastAxis), context.SetConditionStyle(common.SubtotalDecimals, 200))
	ctx.NewLine()
	ctx.SetStringLine("test", false)
	ctx.NewLine()
	ctx.SetStringLine("备注:xxxxx", false)
	ctx.NewLine()
	fmt.Println(ctx.GetValue("A2"))
	fmt.Println(ctx)
	return nil
}
