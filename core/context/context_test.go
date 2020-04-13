package context

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/bluebuff/simple-excelize/core/common"
	"github.com/bluebuff/simple-excelize/models"
	"testing"
)

func TestNewContext(t *testing.T) {
	file := excelize.NewFile()
	file.NewSheet("列表")
	styleManager := common.NewStyleMng()
	common.LoadDefaultStyle(styleManager, file)
	ctx := NewContext(file, styleManager, "列表", models.DefaultLayout)
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
	ctx.SetHeader("小计")
	ctx.Jump()
	ctx.SetFormula(fmt.Sprintf("=SUM(%s:%s)", firstAxis, lastAxis), SetConditionStyle(common.SubtotalDecimals, 200))
	file.SaveAs("./test.xlsx")
}
