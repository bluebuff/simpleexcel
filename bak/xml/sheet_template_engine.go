package xml

import "github.com/bluebuff/simple-excelize/core/context"

type SheetTemplateEngine struct {
	templateManager *TemplateManager
	Error           error
	schema          string
}

func Open(filePattern string) *SheetTemplateEngine {
	templateManager := NewTemplateManager()
	p := NewParse(templateManager)
	err := p.Parse(filePattern)
	if err != nil {
		panic(err)
	}
	return &SheetTemplateEngine{templateManager: templateManager}
}

func (engine *SheetTemplateEngine) Schema(schema string) *SheetTemplateEngine {
	e := engine.clone()
	e.schema = schema
	return e
}

func (engine *SheetTemplateEngine) Scan(data interface{}) ([]context.Handler, error) {
	return engine.newScope(data).Scan()
}

func (engine *SheetTemplateEngine) Print() {
	engine.newScope(nil).Print()
}

func (engine *SheetTemplateEngine) newScope(data interface{}) *SheetScope {
	e := engine.clone()
	return &SheetScope{Data: data, id: e.schema, templateManager: e.templateManager}
}

func (engine *SheetTemplateEngine) clone() *SheetTemplateEngine {
	e := &SheetTemplateEngine{
		templateManager: engine.templateManager,
		schema:          engine.schema,
		Error:           engine.Error,
	}
	return e
}
