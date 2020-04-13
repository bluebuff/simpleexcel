package xml

import (
	"fmt"
	"github.com/beevik/etree"
	"github.com/bluebuff/simple-excelize/core/common"
	"github.com/bluebuff/simple-excelize/core/context"
	"github.com/bluebuff/simple-excelize/util"
	"strings"
)

type SheetParse struct {
	data string
}

func newSheetParse(data string) *SheetParse {
	return &SheetParse{data: data}
}

func (parse *SheetParse) SmartHandle(ctx context.Context) error {
	doc := etree.NewDocument()
	if err := doc.ReadFromString(parse.data); err != nil {
		return err
	}
	root := doc.SelectElement("table")
	rowLevelMap := make(map[int]int)
	for i, tr := range root.SelectElements("tr") {
		colIndex := ctx.CurrentColCount()
		level := rowLevelMap[colIndex]
		if level > 0 {
			ctx.Jump()
		}
		for _, child := range tr.ChildElements() {
			rowspan := util.MustInt(child.SelectAttrValue("rowspan", "0"))
			if rowspan > 0 {
				jumpColIndex := ctx.CurrentColCount()
				rowLevelMap[jumpColIndex] = rowspan
			}
			handleTitleElement(child, ctx, i+1)
		}
		if rowLevelMap[colIndex] > 0 {
			rowLevelMap[colIndex] = rowLevelMap[colIndex] - 1
		}
		ctx.NewLine()
	}
	return nil
}

func handleTitleElement(child *etree.Element, ctx context.Context, rowCount int) {
	colspan := util.MustInt(child.SelectAttrValue("colspan", "1")) - 1
	rowspan := util.MustInt(child.SelectAttrValue("rowspan", "1")) - 1
	t := Type(child.SelectAttrValue("type", ""))
	formula := FormulaFormatAlias(child.SelectAttrValue("formula", ""))
	text := strings.Replace(child.Text(), "$rowCount", fmt.Sprintf("%d", rowCount), -1)
	startAxis := ctx.CurrentAxis()
	if t == "" && child.Tag == "th" {
		t = Header
	}
	format := FormatMap[formula]
	setValue(t, ctx, handleFormat(text, format))
	endAxis := ctx.GetAxis(colspan-1, rowspan)
	ctx.MergeValue(startAxis, endAxis)
}

func handleFormat(text, format string) string {
	if format == "" {
		return text
	}
	axisArray := strings.Split(text, ":")
	startAxis, err := parseAxis(axisArray[0])
	if err != nil {
		return ""
	}
	endAxis, err := parseAxis(axisArray[1])
	if err != nil {
		return ""
	}
	return fmt.Sprintf(format, startAxis, endAxis)
}

func parseAxis(axisExpress string) (*common.Axis, error) {
	axis := new(common.Axis)
	expressArray := strings.Split(axisExpress, "@")
	col, err := util.Process(expressArray[0])
	if err != nil {
		return nil, err
	}
	row, err := util.Process(expressArray[1])
	if err != nil {
		return nil, err
	}
	axis.Col = int(col)
	axis.Row = int(row)
	return axis, nil
}

func setValue(t Type, ctx context.Context, value string) {
	switch t {
	case Uint32:
		ctx.SetUint32(util.MustUint32(value))
	case Uint64:
		ctx.SetUint64(util.MustUint64(value))
	case Float32:
		ctx.SetFloat32(util.MustFloat32(value))
	case Float64:
		ctx.SetFloat64(util.MustFloat64(value))
	case Header:
		ctx.SetHeader(value)
	case Title:
		ctx.SetTitle(value)
	case Formula:
		// TODO impl compare value
		ctx.SetFormula(value, context.SetConditionStyle(common.SubtotalDecimals, 0))
	default:
		ctx.SetString(value)
	}
}
