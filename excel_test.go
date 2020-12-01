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
	Test string `json:"test" excelName:"测试1"`
	Test2 string `json:"test2" excelName:"测试2"`
	Test3 string `json:"test3" excelName:"测试3"`
	Test4 string `json:"test4" excelName:"测试4"`
	Test5 string `json:"test5" excelName:"测试5"`
	Test6 string `json:"test6" excelName:"测试6"`
	Test7 string `json:"test7" excelName:"测试7"`
	Test8 string `json:"test8" excelName:"测试8"`
	Test9 string `json:"test9" excelName:"测试9"`
	Test15 string `json:"test15" excelName:"测试15"`
	Test16 string `json:"test16" excelName:"测试16"`
	Test17 string `json:"test17" excelName:"测试17"`
	Test18 string `json:"test18" excelName:"测试18"`
	Test19 string `json:"test19" excelName:"测试19"`
}
func TestExcel(t *testing.T) {
	var stru []Resume
	info := Resume{
		Name: "张三",
		Sex:  "男",
		Like: "game",
		Age:  19,
		Test: "1",
		Test2: "1",
		Test3: "1",
		Test4: "1",
		Test5: "1",
		Test6: "11fewfeeww2efefefwefewfefewfewfwefewfe",
		Test7: "1",
		Test8: "1",
		Test9: "11fewfeeww2efefefwefewfefewfewfwefewfe",
		Test15: "1",
		Test16: "11fewfeeww2efefefwefewfefewfewfwefewfe",
		Test17: "11fewfeeww2efefefwefewfefewfewfwefewfe",
		Test18: "1",
		Test19: "11fewfeeww2efefefwefewfefewfewfwefewfe",
	}

	for i := 0; i < 10000;i ++ {
		stru = append(stru,info)
	}
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