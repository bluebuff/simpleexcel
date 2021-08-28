package simpleexcel

import (
	"fmt"
	"github.com/xuri/excelize/v2"
)

func main() {
	f := excelize.NewFile()
	const sheet = "Sheet1"

	if err := f.SetPageMargins(sheet,
		excelize.PageMarginBottom(1.0),
		excelize.PageMarginFooter(1.0),
		excelize.PageMarginHeader(1.0),
		excelize.PageMarginLeft(1.0),
		excelize.PageMarginRight(1.0),
		excelize.PageMarginTop(1.0),
	); err != nil {
		fmt.Println(err)
	}
	f.SetHeaderFooter("Sheet1", &excelize.FormatHeaderFooter{
		DifferentFirst:   true,
		DifferentOddEven: true,
		OddHeader:        "&R&P",
		OddFooter:        "&C&F",
		EvenHeader:       "&L&P",
		EvenFooter:       "&L&D&R&T",
		FirstHeader:      `&CCenter &"-,Bold"Bold&"-,Regular"HeaderU+000A&D`,
	})
	f.SaveAs("./test.xlsx")
}
