package main

import (
	"encoding/json"
)

type Student struct {
	Name  string
	Age   uint32
	Score float64
}

func (stu Student) String() string {
	bytes, _ := json.MarshalIndent(stu, "=>", " ")
	return string(bytes)
}

type StudentData struct {
	Date     string
	Students []*Student
}

func (data StudentData) String() string {
	bytes, _ := json.MarshalIndent(data, " ", " ")
	return string(bytes)
}
