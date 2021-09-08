package simpleexcel

import (
	"github.com/bluebuff/simpleexcel/v2/context"
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
	Build() (string, error)
	WriteTo(w io.Writer) (int64, error)
}

func NewExcelBuilder(m mode, opts ...func(style.StyleManager)) ExcelBuilder {
	file := excelize.NewFile()
	styleManager := style.NewStyleManager(file, opts...)
	factory := newSimpleFactory(file, styleManager)
	return &excelBuilder{
		file:           file,
		sheetNames:     make([]string, 0, 5),
		sheetHandles:   make(map[string][]context.Handler, 5),
		beforeHandlers: make([]context.Handler, 0),
		afterHandlers:  make([]context.Handler, 0),
		StyleManager:   styleManager,
		factory:        factory,
		mode:           m,
	}
}

type excelBuilder struct {
	file            *excelize.File
	sheetNames      []string
	beforeHandlers  []context.Handler
	sheetHandles    map[string][]context.Handler
	afterHandlers   []context.Handler
	activeSheetName string
	StyleManager    style.StyleManager
	factory         Factory
	mode            mode
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

func (builder *excelBuilder) WriteTo(w io.Writer) (int64, error) {
	if err := builder.build(); err != nil {
		return 0, err
	}
	return builder.file.WriteTo(w)
}

func (builder *excelBuilder) Build() (string, error) {
	// build
	if err := builder.build(); err != nil {
		return "", err
	}
	// write temp file
	tmp, err := ioutil.TempFile(os.TempDir(), TempFilePattern)
	defer tmp.Close()
	if err != nil {
		return "", err
	}
	if _, err := builder.file.WriteTo(tmp); err != nil {
		return "", err
	}
	return tmp.Name(), nil
}

func (builder *excelBuilder) build() error {
	for i, sheetName := range builder.sheetNames {
		if i == 0 {
			builder.file.SetSheetName("Sheet1", sheetName)
		} else {
			builder.file.NewSheet(sheetName)
		}
		// new context
		ctx, done, err := builder.factory.New(sheetName, builder.mode)
		if err != nil {
			return err
		}
		// before
		if builder.beforeHandlers != nil && len(builder.beforeHandlers) != 0 {
			for _, handler := range builder.beforeHandlers {
				if err := handler(ctx); err != nil {
					return err
				}
			}
		}
		for _, handler := range builder.sheetHandles[sheetName] {
			if err := handler(ctx); err != nil {
				return err
			}
		}
		// after
		if builder.afterHandlers != nil && len(builder.afterHandlers) != 0 {
			for _, handler := range builder.afterHandlers {
				if err := handler(ctx); err != nil {
					return err
				}
			}
		}
		if err := done(); err != nil {
			return err
		}
	}

	index := builder.file.GetSheetIndex(builder.activeSheetName)
	builder.file.SetActiveSheet(index)
	return nil
}

func (builder *excelBuilder) Before(handle context.Handler) {
	builder.beforeHandlers = append(builder.beforeHandlers, handle)
}

func (builder *excelBuilder) After(handle context.Handler) {
	builder.afterHandlers = append(builder.afterHandlers, handle)
}
