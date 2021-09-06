package simpleexcel

import (
	"github.com/bluebuff/simpleexcel/v2/context"
	"github.com/bluebuff/simpleexcel/v2/context/normal"
	"github.com/bluebuff/simpleexcel/v2/context/streamwriter"
	"github.com/bluebuff/simpleexcel/v2/style"
	"github.com/xuri/excelize/v2"
)

type mode uint32

const (
	Normal      mode = 1
	StreamWrite mode = 2
)

type Factory interface {
	New(sheetName string, m mode) (context.Context, func() error, error)
}

type simpleFactory struct {
	file         *excelize.File
	styleManager style.StyleManager
}

func newSimpleFactory(file *excelize.File, styleManager style.StyleManager) Factory {
	return &simpleFactory{
		file:         file,
		styleManager: styleManager,
	}
}

func (s *simpleFactory) New(sheetName string, m mode) (ctx context.Context, done func() error, err error) {
	switch m {
	case Normal:
		ctx = normal.NewContext(s.file, s.styleManager, sheetName)
		done = func() error { return nil }
	case StreamWrite:
		sw, err := s.file.NewStreamWriter(sheetName)
		if err != nil {
			return nil, nil, err
		}
		ctx = streamwriter.NewContext(sw, s.styleManager)
		done = func() error { return sw.Flush() }
	}
	return
}
