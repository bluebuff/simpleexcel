package main

import (
	"fmt"
	"github.com/bluebuff/simpleexcel/v2"
	"github.com/bluebuff/simpleexcel/v2/context"
	"github.com/bluebuff/simpleexcel/v2/context/streamwriter"
	"github.com/bluebuff/simpleexcel/v2/style/standard"
	"github.com/xuri/excelize/v2"
	"math/rand"
)

func main() {
	file := excelize.NewFile()

	// style
	styleManager := simpleexcel.NewStyleManager(file).Configure(standard.Style)

	// sheet1
	sw, err := file.NewStreamWriter("Sheet1")
	if err != nil {
		fmt.Println(err)
		return
	}

	ctx := streamwriter.NewContext(sw, styleManager)
	do(ctx)
	sw.Flush()

	// sheet2
	file.NewSheet("Sheet2")
	sw, err = file.NewStreamWriter("Sheet2")
	if err != nil {
		fmt.Println(err)
		return
	}
	ctx = streamwriter.NewContext(sw, styleManager)
	do2(ctx)
	sw.Flush()

	// save
	file.SaveAs("./test.xlsx")
}

func do2(ctx context.Context) {

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
}

func do(ctx context.Context) {
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
	ctx.SetUint32(28, simpleexcel.CompareLessAndStyleUint32(30))
	ctx.SetFloat32(-99, simpleexcel.CompareLessAndStyleFloat32(0))
	lastAxis := ctx.LastAxis()
	ctx.NewLine()
	ctx.SetStringLine("")
	ctx.NewLine()
	d := ctx.SetHeader("小计")
	fmt.Println(d)
	ctx.SetHeader("")
	f := ctx.SetFormula(fmt.Sprintf("=SUM(%s:%s)", firstAxis, lastAxis), simpleexcel.CustomStyleByAlias(simpleexcel.SubtotalDecimalsCondition))
	fmt.Println(f)
	ctx.NewLine()
	g := ctx.SetStringLine("test")
	fmt.Println(g)
	ctx.NewLine()
	hcell := ctx.SetString("Cell1")
	ctx.SetString("Cell1-")
	ctx.SetString("Cell1--")
	ctx.NewLine()
	vcell := ctx.SetString("Cell2")
	ctx.SetString("Cell2-")
	ctx.SetString("Cell2--")
	ctx.NewLine()
	ctx.MergeValue(hcell, vcell)
}
