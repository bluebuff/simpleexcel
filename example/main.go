package main

import (
	"github.com/bluebuff/simple-excelize/core"
	"github.com/bluebuff/simple-excelize/xml"
	"io/ioutil"
	"time"
)

func main() {
	engine := xml.Open("./config/student.xml")
	studentDataList := NewStudentDataList()
	builder := core.NewExcelBuilder()
	builder.RegisterStyle()
	handlers, err := engine.Schema("student-list").Scan(studentDataList)
	if err != nil {
		panic(err)
	}
	builder.JoinSheet("列表1", handlers...)
	bytes, err := builder.Build()
	if err != nil {
		panic(err)
	}
	err = ioutil.WriteFile("./test1.xlsx", bytes, 0666)
	if err != nil {
		panic(err)
	}
	handlers, err = engine.Schema("st").Scan(studentDataList)
	if err != nil {
		panic(err)
	}
	bytes, err = builder.Build()
	if err != nil {
		panic(err)
	}
	err = ioutil.WriteFile("./test2.xlsx", bytes, 0666)
	if err != nil {
		panic(err)
	}
}

func NewStudentDataList() []*StudentData {
	studentDataList := make([]*StudentData, 0)
	now := time.Now()
	studentDataList = append(studentDataList, &StudentData{
		Date:     now.Add(time.Hour * 24 * 1).Format("2006-01-02"),
		Students: NewStudents(100),
	})
	studentDataList = append(studentDataList, &StudentData{
		Date:     now.Add(time.Hour * 24 * 2).Format("2006-01-02"),
		Students: NewStudents(200),
	})
	studentDataList = append(studentDataList, &StudentData{
		Date:     now.Add(time.Hour * 24 * 3).Format("2006-01-02"),
		Students: NewStudents(300),
	})
	return studentDataList
}

func NewStudents(score float64) []*Student {
	students := make([]*Student, 0)
	students = append(students, &Student{
		Name:  "张三",
		Age:   20,
		Score: score + 10,
	})
	students = append(students, &Student{
		Name:  "李四",
		Age:   25,
		Score: score + 20,
	})
	students = append(students, &Student{
		Name:  "王五",
		Age:   26,
		Score: score + 30,
	})
	return students
}
