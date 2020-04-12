package xml

import (
	"errors"
	"github.com/beevik/etree"
	"github.com/bluebuff/simple-excelize/v1/models"
	"github.com/bluebuff/simple-excelize/v1/util"
	"strconv"
	"strings"
	"text/template"
)

// error
var file_not_exist = errors.New("the file not exist")

type Parser interface {
	Parse(fileName string) error
}

type TemplateParser struct {
	templateManager *TemplateManager
}

func NewParse(templateManager *TemplateManager) Parser {
	parse := &TemplateParser{
		templateManager: templateManager,
	}
	return parse
}

func (parse *TemplateParser) Parse(fileName string) error {
	doc := etree.NewDocument()
	if err := doc.ReadFromFile(fileName); err != nil {
		return err
	}
	root := doc.SelectElement("template")
	for _, ele := range root.ChildElements() {
		var err error
		switch strings.ToLower(ele.Tag) {
		case "style":
			err = parse.parseStyle(ele)
		case "schema":
			err = parse.parseSchema(ele)
		default:
			return errors.New("not support element:" + ele.Tag)
		}
		if err != nil {
			return err
		}
	}
	return nil
}

func (parse *TemplateParser) parseStyle(element *etree.Element) error {
	m := element.SelectElement("map")
	if m == nil {
		return errors.New("style map is empty")
	}
	for i, entry := range m.SelectElements("entry") {
		text := strings.TrimSpace(entry.Text())
		if text == "" {
			continue
		}
		key := strconv.Itoa(i) // key default is index
		attr := entry.SelectAttr("key")
		if attr != nil {
			key = attr.Value
		}
		parse.templateManager.SetStyle(key, text)
	}
	return nil
}

func (parse *TemplateParser) parseSchema(element *etree.Element) error {
	schema := &models.Schema{}
	// parse id
	idEle := element.SelectElement("id")
	if idEle == nil || strings.TrimSpace(idEle.Text()) == "" {
		return errors.New("id is empty")
	}
	schema.ID = strings.TrimSpace(idEle.Text())
	canvasEle := element.SelectElement("canvas")
	if canvasEle != nil {
		widths := make([]float64, 0)
		if widthsEle := canvasEle.SelectElement("widths"); widthsEle != nil {
			for _, widthEle := range widthsEle.SelectElements("width") {
				widths = append(widths, util.MustFloat64(widthEle.Text()))
			}
		}
		schema.Widths = widths
		// parse layout [option]
		layoutEle := canvasEle.SelectElement("layout")
		schema.Layout = models.DefaultLayout
		if layoutEle != nil {
			leftEle := layoutEle.SelectElement("left")
			if leftEle != nil && leftEle.Text() != "" {
				schema.Layout.Left = util.MustInt(leftEle.Text())
			}
			topEle := layoutEle.SelectElement("top")
			if topEle != nil && topEle.Text() != "" {
				schema.Layout.Top = util.MustInt(topEle.Text())
			}
			rightEle := layoutEle.SelectElement("right")
			if rightEle != nil && rightEle.Text() != "" {
				schema.Layout.Right = util.MustInt(rightEle.Text())
			}
			bottomEle := layoutEle.SelectElement("bottom")
			if bottomEle != nil && bottomEle.Text() != "" {
				schema.Layout.Bottom = util.MustInt(bottomEle.Text())
			}
		}
	}
	// parse sheet template
	sheetEle := element.SelectElement("sheet")
	if sheetEle == nil {
		return errors.New("sheet not exist!")
	}
	schema.Sheet = strings.TrimSpace(sheetEle.Text())
	// new template
	tpl, err := template.New(schema.ID).Parse(schema.Sheet)
	if err != nil {
		return err
	}
	parse.templateManager.SetTemplate(schema.ID, tpl)
	parse.templateManager.SetWidths(schema.ID, schema.Widths)
	return nil
}
