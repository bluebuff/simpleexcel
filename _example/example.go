package main

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"math/rand"
)

func main() {
	file := excelize.NewFile()

	file.SetColWidth("Sheet1", "A", "E", 20)

	streamWriter, err := file.NewStreamWriter("Sheet1")
	if err != nil {
		fmt.Println(err)
		return
	}

	streamWriter.SetColWidth(1, 50, 20)

	headers := make([]interface{}, 0, 50)

	styleId, _ := file.NewStyle(`{"fill":{"type":"pattern","pattern":1,"color":["#E8E8E8"]},"font":{"bold":true,"italic":false,"family":"正楷","size":12,"color":"#000000"},"alignment":{"horizontal":"center","vertical":"center"},"border":[{"type":"left","color":"#000000","style":1},{"type":"top","color":"#000000","style":1},{"type":"right","color":"#000000","style":1},{"type":"bottom","color":"#000000","style":1}]}`)

	for i := 0; i < 50; i++ {
		headers = append(headers, excelize.Cell{
			StyleID: styleId,
			Formula: "",
			Value:   fmt.Sprintf("列%d", i+1),
		})
	}

	streamWriter.SetRow("A1", headers)

	styleId2, _ := file.NewStyle(`{"font":{"bold":false,"italic":false,"family":"正楷","size":12,"color":"#000000"},"alignment":{"horizontal":"center","vertical":"center"},"border":[{"type":"left","color":"#000000","style":1},{"type":"top","color":"#000000","style":1},{"type":"right","color":"#000000","style":1},{"type":"bottom","color":"#000000","style":1}]}`)
	for row := 2; row <= 500; row++ {
		values := make([]interface{}, 0, 50)
		for i := 0; i < 50; i++ {
			values = append(values, excelize.Cell{
				StyleID: styleId2,
				Formula: "",
				Value:   fmt.Sprintf("%d", rand.Int()),
			})
		}
		streamWriter.SetRow(fmt.Sprintf("A%d", row), values)
		if (row-1)%5 == 0 {
			streamWriter.MergeCell(fmt.Sprintf("A%d", row-4), fmt.Sprintf("A%d", row))
		}
	}

	streamWriter.Flush()

	file.SaveAs("./test.xlsx")
}
