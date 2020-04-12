package xml

import (
	"bytes"
	"fmt"
	"github.com/bluebuff/simple-excelize/v1/core/context"
)

type SheetScope struct {
	id              string
	Data            interface{}
	Value           string
	templateManager *TemplateManager
}

func (scope *SheetScope) Scan() ([]context.Handler, error) {
	template := scope.templateManager.GetTemplate(scope.id)
	if template == nil {
		return nil, fmt.Errorf("The template '%s' not found!", scope.id)
	}
	var buff bytes.Buffer
	if err := template.Execute(&buff, scope.Data); err != nil {
		return nil, err
	}
	tableParse := newSheetParse(buff.String())
	handler := make([]context.Handler, 0, 2)
	// before handle
	handler = append(handler, beforeHandle(scope.templateManager, scope.id))
	// template handle
	handler = append(handler, tableParse.SmartHandle)
	return handler, nil
}

func beforeHandle(templateManager *TemplateManager, id string) context.Handler {
	layout := templateManager.GetLayout(id)
	widths := templateManager.GetWidths(id)
	return func(ctx context.Context) error {
		ctx.SetLayout(layout)
		for i, width := range widths {
			ctx.SetColWidth(i+1, i+1, width)
		}
		return nil
	}
}

func (scope *SheetScope) Print() {
	fmt.Println(scope.Value)
}
