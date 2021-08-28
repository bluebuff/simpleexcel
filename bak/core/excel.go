package core

import (
	"github.com/bluebuff/simple-excelize/core/common"
	"github.com/bluebuff/simple-excelize/core/context"
	"github.com/bluebuff/simple-excelize/models"
	"github.com/xuri/excelize/v2"
)

type ExcelBuilder interface {
	JoinSheet(sheetName string, handle ...context.Handler)
	Active(sheetName string)
	Before(handle context.Handler)
	After(handle context.Handler)
	RegisterStyle(customStyleFunc ...func(*common.StyleManager))
	Build() ([]byte, error)
}

func NewExcelBuilder() ExcelBuilder {
	builder := &excelBuilder{
		file:           excelize.NewFile(),
		sheetNames:     make([]string, 0, 5),
		sheetHandles:   make(map[string][]context.Handler, 5),
		beforeHandlers: make([]context.Handler, 0),
		afterHandlers:  make([]context.Handler, 0),
		styleManager:   common.NewStyleMng(),
	}
	return builder
}

type excelBuilder struct {
	file            *excelize.File
	sheetNames      []string
	beforeHandlers  []context.Handler
	sheetHandles    map[string][]context.Handler
	afterHandlers   []context.Handler
	styleManager    *common.StyleManager
	activeSheetName string
	ExcelBuilder
}

func (builder *excelBuilder) Active(sheetName string) {
	builder.activeSheetName = sheetName
}

func (builder *excelBuilder) JoinSheet(sheetName string, handler ...context.Handler) {
	if _, ok := builder.sheetHandles[sheetName]; !ok {
		builder.sheetNames = append(builder.sheetNames, sheetName)
	}
	builder.sheetHandles[sheetName] = append(builder.sheetHandles[sheetName], handler...)
}

func (builder *excelBuilder) Build() ([]byte, error) {

	for i, sheetName := range builder.sheetNames {
		ctx := context.NewContext(builder.file, builder.styleManager, sheetName, models.DefaultLayout)
		if i == 0 {
			builder.file.SetSheetName("Sheet1", sheetName)
		} else {
			builder.file.NewSheet(sheetName)
		}
		// before
		if builder.beforeHandlers != nil && len(builder.beforeHandlers) != 0 {
			for _, handler := range builder.beforeHandlers {
				if err := handler(ctx); err != nil {
					return nil, err
				}
			}
		}
		for _, handler := range builder.sheetHandles[sheetName] {
			if err := handler(ctx); err != nil {
				return nil, err
			}
		}
		// after
		if builder.afterHandlers != nil && len(builder.afterHandlers) != 0 {
			for _, handler := range builder.afterHandlers {
				if err := handler(ctx); err != nil {
					return nil, err
				}
			}
		}
	}

	index := builder.file.GetSheetIndex(builder.activeSheetName)
	builder.file.SetActiveSheet(index)

	buff, err := builder.file.WriteToBuffer()
	if err != nil {
		return nil, err
	}
	return buff.Bytes(), nil
}

func (builder *excelBuilder) Before(handle context.Handler) {
	builder.beforeHandlers = append(builder.beforeHandlers, handle)
}

func (builder *excelBuilder) After(handle context.Handler) {
	builder.afterHandlers = append(builder.afterHandlers, handle)
}

func (builder *excelBuilder) RegisterStyle(customStyleFunc ...func(*common.StyleManager)) {
	// default style
	common.LoadDefaultStyle(builder.styleManager, builder.file)
	// custom style
	if customStyleFunc != nil && len(customStyleFunc) != 0 {
		customStyleFunc[0](builder.styleManager)
	}
}
