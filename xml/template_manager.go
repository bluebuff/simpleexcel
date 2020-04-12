package xml

import (
	"github.com/bluebuff/simple-excelize/v1/models"
	"sync"
	"text/template"
)

type TemplateManager struct {
	templateMap map[string]*template.Template
	styleMap    map[string]string
	layoutMap   map[string]*models.Layout
	widthsMap   map[string][]float64
	sync        sync.RWMutex
}

func NewTemplateManager() *TemplateManager {
	mng := &TemplateManager{
		templateMap: make(map[string]*template.Template),
		styleMap:    make(map[string]string),
		layoutMap:   make(map[string]*models.Layout),
		widthsMap:   make(map[string][]float64),
	}
	return mng
}

func (mng *TemplateManager) GetTemplate(id string) *template.Template {
	mng.sync.RLock()
	defer mng.sync.RUnlock()
	return mng.templateMap[id]
}

func (mng *TemplateManager) SetTemplate(id string, tpl *template.Template) *template.Template {
	old := mng.templateMap[id]
	mng.sync.Lock()
	mng.templateMap[id] = tpl
	mng.sync.Unlock()
	return old
}

func (mng *TemplateManager) SetStyle(key, style string) string {
	old := mng.styleMap[key]
	mng.styleMap[key] = style
	return old
}

func (mng *TemplateManager) GetStyle(key string) string {
	return mng.styleMap[key]
}

func (mng *TemplateManager) SetLayout(id string, layout *models.Layout) *models.Layout {
	old := mng.layoutMap[id]
	mng.layoutMap[id] = layout
	return old
}

func (mng *TemplateManager) GetLayout(id string) *models.Layout {
	return mng.layoutMap[id]
}

func (mng *TemplateManager) SetWidths(id string, widths []float64) []float64 {
	old := mng.widthsMap[id]
	mng.widthsMap[id] = widths
	return old
}

func (mng *TemplateManager) GetWidths(id string) []float64 {
	return mng.widthsMap[id]
}
