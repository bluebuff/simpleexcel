package models

import "encoding/json"

type Schema struct {
	ID     string
	Widths []float64
	Layout Layout
	Sheet  string
}

func (schema Schema) String() string {
	data, _ := json.Marshal(schema)
	return string(data)
}
