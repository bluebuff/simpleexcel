package core

import (
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/bluebuff/simple-excelize/v1/core/common"
	"github.com/bluebuff/simple-excelize/v1/core/context"
	"github.com/bluebuff/simple-excelize/v1/models"
)

type ExcelBuilder interface {
	JoinSheet(sheetName string, handle ...context.Handler)
	Active(sheetName string)
	RegisterStyle(customStyleFunc ...func(*common.StyleManager))
	Build() ([]byte, error)
}

func NewExcelBuilder() ExcelBuilder {
	builder := &excelBuilder{
		file:         excelize.NewFile(),
		sheetNames:   make([]string, 0, 5),
		sheetHandles: make(map[string][]context.Handler, 5),
		styleManager: common.NewStyleMng(),
	}
	return builder
}

type excelBuilder struct {
	file         *excelize.File
	sheetNames   []string
	sheetHandles map[string][]context.Handler
	styleManager *common.StyleManager
	ExcelBuilder
}

func (builder *excelBuilder) Active(sheetName string) {
	index := builder.file.GetSheetIndex(sheetName)
	builder.file.SetActiveSheet(index)
}

func (builder *excelBuilder) JoinSheet(sheetName string, handler ...context.Handler) {
	builder.sheetNames = append(builder.sheetNames, sheetName)
	builder.sheetHandles[sheetName] = handler
}

func (builder *excelBuilder) Build() ([]byte, error) {
	for i, sheetName := range builder.sheetNames {
		ctx := context.NewContext(builder.file, builder.styleManager, sheetName, models.DefaultLayout)
		if i == 0 {
			builder.file.SetSheetName("Sheet1", sheetName)
		} else {
			builder.file.NewSheet(sheetName)
		}
		for _, handler := range builder.sheetHandles[sheetName] {
			if err := handler(ctx); err != nil {
				return nil, err
			}
		}
	}
	buff, err := builder.file.WriteToBuffer()
	if err != nil {
		return nil, err
	}
	return buff.Bytes(), nil
}

func (builder *excelBuilder) RegisterStyle(customStyleFunc ...func(*common.StyleManager)) {
	// default style
	common.LoadDefaultStyle(builder.styleManager, builder.file)
	// custom style
	if customStyleFunc != nil && len(customStyleFunc) != 0 {
		customStyleFunc[0](builder.styleManager)
	}
}
