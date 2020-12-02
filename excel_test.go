package excel

import (
	"fmt"
	"testing"
	"time"
)

type Resume struct {
	Name string `json:"name" excel_name:"我的名字"`
	Like string `json:"like"`
	Sex string `json:"sex" excel_name:"我的性别"`
	Age int `json:"age" excel_name:"我的年龄"`
	Status int8 `json:"status" excel_name:"当前状态" enums:"0:冻结,1:开启,2:关闭"` //如设置enums 数据类型必须为 int8
	CreateTime int64 `json:"create_time" excel_name:"创建时间" excel_time:"int"` //如设置 excel_time为int 数据类型必须为 int64
	UpdateTime time.Time `json:"update_time" excel_name:"更新时间" excel_time:"time"`
	Test string `json:"test" excel_name:"测试1"`
	Test2 string `json:"test2" excel_name:"测试2"`
	Test3 string `json:"test3" excel_name:"测试3"`
	Test4 string `json:"test4" excel_name:"测试4"`
	Test5 string `json:"test5" excel_name:"测试5"`
	Test6 string `json:"test6" excel_name:"测试6"`
	Test7 string `json:"test7" excel_name:"测试7"`
	Test8 string `json:"test8" excel_name:"测试8"`
	Test9 string `json:"test9" excel_name:"测试9"`
	Test15 string `json:"test15" excel_name:"测试15"`
	Test16 string `json:"test16" excel_name:"测试16"`
	Test17 string `json:"test17" excel_name:"测试17"`
	Test18 string `json:"test18" excel_name:"测试18"`
	Test19 string `json:"test19" excel_name:"测试19"`
}
func TestExcel(t *testing.T) {
	var stru []Resume
	info := Resume{
		Name: "张三",
		Sex:  "男",
		Like: "game",
		Age:  19,
		CreateTime: time.Now().Unix(),
		UpdateTime: time.Now(),
		Status: 2,
		Test: "1",
		Test2: "1",
		Test3: "1",
		Test4: "1",
		Test5: "1",
		Test6: "1",
		Test7: "1",
		Test8: "1",
		Test9: "1",
		Test15: "1",
		Test16: "1",
		Test17: "1",
		Test18: "1",
		Test19: "1",
	}

	for i := 0; i < 10;i ++ {
		stru = append(stru,info)
	}
	e := Excel{
		Sheet: "Sheet1",
	}
	err := e.SaveExcel("test.xlsx",&stru)
	fmt.Println(err)
	return
}
func TestExcel_ReadExcel(t *testing.T) {
	e := Excel{
		Sheet: "Sheet1",
	}
	var stru []*Resume
	err := e.ReadExcel("test.xlsx",&stru)
	fmt.Println(err)
	for _,v := range stru {
		fmt.Println(v)
	}
	return
}
func TestAddStr(t *testing.T){

	fmt.Println(addStr("BZ",3))
}
