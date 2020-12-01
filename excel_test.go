package excel

import (
	"fmt"
	"testing"
)

type Resume struct {
	Name string `json:"name" excelName:"我的名字"`
	Like string `json:"like"`
	Sex string `json:"sex" excelName:"我的性别"`
	Age int `json:"age" excelName:"我的年龄"`
}
func TestExcel(t *testing.T) {
	var stru []Resume
	info := Resume{
		Name: "张三",
		Sex:  "男",
		Like: "game",
		Age:  19,
	}
	stru = append(stru,info)
	info2 := Resume{
		Sex:  "男",
		Like: "car",
		Age:  20,
		Name: "李四",
	}
	stru = append(stru,info2)
	e := Excel{}
	err := e.SaveExcel("test.xlsx",&stru)
	fmt.Println(err)
	return
}
func TestExcel_ReadExcel(t *testing.T) {
	e := Excel{}
	var stru []Resume
	err := e.ReadExcel("test.xlsx",&stru)
	fmt.Println(err)
	fmt.Println(stru)
	return
}