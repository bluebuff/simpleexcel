package context

import (
	"github.com/stretchr/testify/assert"
	"github.com/xuri/excelize/v2"
	"math/rand"
	"path/filepath"
	"strings"
	"testing"
	"time"
)

func TestNewContext(t *testing.T) {
	file := excelize.NewFile()
	streamWriter, err := file.NewStreamWriter("Sheet1")
	assert.NoError(t, err)

	// Test max characters in a cell.
	row := make([]interface{}, 1)
	row[0] = strings.Repeat("c", excelize.TotalCellChars+2)
	assert.NoError(t, streamWriter.SetRow("A1", row))

	// Test leading and ending space(s) character characters in a cell.
	row = make([]interface{}, 1)
	row[0] = " characters"
	assert.NoError(t, streamWriter.SetRow("A2", row))

	row = make([]interface{}, 1)
	row[0] = []byte("Word")
	assert.NoError(t, streamWriter.SetRow("A3", row))

	// Test set cell with style.
	styleID, err := file.NewStyle(`{"font":{"color":"#777777"}}`)
	assert.NoError(t, err)
	assert.NoError(t, streamWriter.SetRow("A4", []interface{}{excelize.Cell{StyleID: styleID}, excelize.Cell{Formula: "SUM(A10,B10)"}}), excelize.RowOpts{Height: 45})
	assert.NoError(t, streamWriter.SetRow("A5", []interface{}{&excelize.Cell{StyleID: styleID, Value: "cell"}, &excelize.Cell{Formula: "SUM(A10,B10)"}}))
	assert.NoError(t, streamWriter.SetRow("A6", []interface{}{time.Now()}))
	assert.NoError(t, streamWriter.SetRow("A7", nil, excelize.RowOpts{Hidden: true}))
	assert.EqualError(t, streamWriter.SetRow("A7", nil, excelize.RowOpts{Height: excelize.MaxRowHeight + 1}), excelize.ErrMaxRowHeight.Error())

	for rowID := 10; rowID <= 51200; rowID++ {
		row := make([]interface{}, 50)
		for colID := 0; colID < 50; colID++ {
			row[colID] = rand.Intn(640000)
		}
		cell, _ := excelize.CoordinatesToCellName(1, rowID)
		assert.NoError(t, streamWriter.SetRow(cell, row))
	}

	assert.NoError(t, streamWriter.Flush())



	// Save spreadsheet by the given path.
	assert.NoError(t, file.SaveAs(filepath.Join("./", "TestStreamWriter.xlsx")))
}


func Test1(t *testing.T) {
	file := excelize.NewFile()
	streamWriter, err := file.NewStreamWriter("Sheet1")
	assert.NoError(t, err)

	// Test max characters in a cell.
	row := make([]interface{}, 1)
	row[0] = strings.Repeat("c", excelize.TotalCellChars+2)
	assert.NoError(t, streamWriter.SetRow("A1", row))

	// Test leading and ending space(s) character characters in a cell.
	row = make([]interface{}, 1)
	row[0] = " characters"
	assert.NoError(t, streamWriter.SetRow("A2", row))

	row = make([]interface{}, 1)
	row[0] = []byte("Word")
	assert.NoError(t, streamWriter.SetRow("A3", row))

	// Test set cell with style.
	styleID, err := file.NewStyle(`{"font":{"color":"#777777"}}`)
	assert.NoError(t, err)
	assert.NoError(t, streamWriter.SetRow("A4", []interface{}{excelize.Cell{StyleID: styleID}, excelize.Cell{Formula: "SUM(A10,B10)"}}), excelize.RowOpts{Height: 45})
	assert.NoError(t, streamWriter.SetRow("A5", []interface{}{&excelize.Cell{StyleID: styleID, Value: "cell"}, &excelize.Cell{Formula: "SUM(A10,B10)"}}))
	assert.NoError(t, streamWriter.SetRow("A6", []interface{}{time.Now()}))
	assert.NoError(t, streamWriter.SetRow("A7", nil, excelize.RowOpts{Hidden: true}))
	assert.EqualError(t, streamWriter.SetRow("A7", nil, excelize.RowOpts{Height: excelize.MaxRowHeight + 1}), excelize.ErrMaxRowHeight.Error())

	t.Log("1")
	assert.NoError(t, streamWriter.Flush())
	t.Log("2")

	file.NewSheet("Sheet2")
	streamWriter, err = file.NewStreamWriter("Sheet2")
	assert.NoError(t, err)

	for rowID := 10; rowID <= 51200; rowID++ {
		row := make([]interface{}, 50)
		for colID := 0; colID < 50; colID++ {
			row[colID] = rand.Intn(640000)
		}
		cell, _ := excelize.CoordinatesToCellName(1, rowID)
		assert.NoError(t, streamWriter.SetRow(cell, row))
	}

	assert.NoError(t, streamWriter.Flush())


	// Save spreadsheet by the given path.
	assert.NoError(t, file.SaveAs(filepath.Join("./", "TestStreamWriter2.xlsx")))
}