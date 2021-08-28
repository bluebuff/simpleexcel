package simpleexcel

import (
	"github.com/bluebuff/simpleexcel/v2/context"
	"github.com/bluebuff/simpleexcel/v2/context/streamwriter"
	"github.com/bluebuff/simpleexcel/v2/style"
	"github.com/xuri/excelize/v2"
	"io"
	"io/ioutil"
	"os"
)

const TempFilePattern = `simple-excel-*.tmp`

type ExcelBuilder interface {
	JoinSheet(sheetName string, handle ...context.Handler)
	Active(sheetName string)
	Before(handle context.Handler)
	After(handle context.Handler)
	Build() (io.Reader, error)
}

func NewStreamWriterExcelBuilder(opts ...func(style.StyleManager)) ExcelBuilder {
	file := excelize.NewFile()
	return &streamWriterExcelBuilder{
		file:           file,
		sheetNames:     make([]string, 0, 5),
		sheetHandles:   make(map[string][]context.Handler, 5),
		beforeHandlers: make([]context.Handler, 0),
		afterHandlers:  make([]context.Handler, 0),
		StyleManager:   style.NewStyleManager(file).Configure(opts...),
	}
}

type streamWriterExcelBuilder struct {
	file            *excelize.File
	sheetNames      []string
	beforeHandlers  []context.Handler
	sheetHandles    map[string][]context.Handler
	afterHandlers   []context.Handler
	activeSheetName string
	StyleManager    style.StyleManager
	ExcelBuilder
}

func (builder *streamWriterExcelBuilder) Active(sheetName string) {
	builder.activeSheetName = sheetName
}

func (builder *streamWriterExcelBuilder) JoinSheet(sheetName string, handler ...context.Handler) {
	if _, ok := builder.sheetHandles[sheetName]; !ok {
		builder.sheetNames = append(builder.sheetNames, sheetName)
	}
	builder.sheetHandles[sheetName] = append(builder.sheetHandles[sheetName], handler...)
}

func (builder *streamWriterExcelBuilder) Build() (io.Reader, error) {

	for i, sheetName := range builder.sheetNames {
		if i == 0 {
			builder.file.SetSheetName("Sheet1", sheetName)
		} else {
			builder.file.NewSheet(sheetName)
		}
		sw, err := builder.file.NewStreamWriter(sheetName)
		if err != nil {
			return nil, err
		}
		// new context
		ctx := streamwriter.NewContext(sw, builder.StyleManager)
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
		if err := sw.Flush(); err != nil {
			return nil, err
		}
	}

	index := builder.file.GetSheetIndex(builder.activeSheetName)
	builder.file.SetActiveSheet(index)

	// write temp file
	tmp, err := ioutil.TempFile(os.TempDir(), TempFilePattern)
	if err != nil {
		return nil, err
	}
	if _, err := builder.file.WriteTo(tmp); err != nil {
		return nil, err
	}
	return tmp, nil
}

func (builder *streamWriterExcelBuilder) Before(handle context.Handler) {
	builder.beforeHandlers = append(builder.beforeHandlers, handle)
}

func (builder *streamWriterExcelBuilder) After(handle context.Handler) {
	builder.afterHandlers = append(builder.afterHandlers, handle)
}
