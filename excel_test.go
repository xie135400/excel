package excel

import (
	"encoding/csv"
	"encoding/json"
	"errors"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"io"
	"os"
	"os/exec"
	"strings"
	"testing"
	"time"
)

type Resume struct {
	Name       string    `json:"name" excel_name:"我的名字"`
	Like       string    `json:"like"`
	Sex        string    `json:"sex" excel_name:""`
	Age        int       `json:"age" excel_name:"我的年龄"`
	Status     int8      `json:"status" excel_name:"当前状态" enums:"0:冻结,1:开启,2:关闭"` //如设置enums 数据类型必须为 int8
	CreateTime int64     `json:"create_time" excel_name:"创建时间" excel_time:"int"`  //如设置 excel_time为int 数据类型必须为 int64
	UpdateTime time.Time `json:"update_time" excel_name:"更新时间" excel_time:"time"`
	Test       string    `json:"test" excel_name:"测试1"`
	Test2      string    `json:"test2" excel_name:"测试2"`
	Test3      string    `json:"test3" excel_name:"测试3"`
	Test4      string    `json:"test4" excel_name:"测试4"`
	Test5      string    `json:"test5" excel_name:"测试5"`
	Test6      string    `json:"test6" excel_name:"测试6"`
	Test7      string    `json:"test7" excel_name:"测试7"`
	Test8      string    `json:"test8" excel_name:"测试8"`
	Test9      string    `json:"test9" excel_name:"测试9"`
	Test15     string    `json:"test15" excel_name:"测试15"`
	Test16     string    `json:"test16" excel_name:"测试16"`
	Test17     string    `json:"test17" excel_name:"测试17"`
	Test18     string    `json:"test18" excel_name:"测试18"`
	Test19     string    `json:"test19" excel_name:"测试19"`
}

type HeatMapData struct {
	Index          string  `json:"index" excel_name:""`
	BuildingsPower float64 `json:"buildings_power" excel_name:"buildings_power"`
	Paid           float64 `json:"paid" excel_name:"paid"`
	Power          float64 `json:"power" excel_name:"power"`
	AllianceId     float64 `json:"alliance_id" excel_name:"alliance_id"`
	Energy         float64 `json:"energy" excel_name:"energy"`
	Language       float64 `json:"language" excel_name:"language"`
	TechsPower     float64 `json:"techs_power" excel_name:"techs_power"`
	PaidTimes      float64 `json:"paid_times" excel_name:"paid_times"`
	BaseLevel      float64 `json:"base_level" excel_name:"base_level"`
	Level          float64 `json:"level" excel_name:"level"`
	BanExpire      float64 `json:"ban_expire" excel_name:"ban_expire"`
	VipLevel       float64 `json:"vip_level" excel_name:"vip_level"`
	Avatar         float64 `json:"avatar" excel_name:"avatar"`
	ArmiesPower    float64 `json:"armies_power" excel_name:"armies_power"`
	IsInternal     float64 `json:"is_internal" excel_name:"is_internal"`
	LoginAt        float64 `json:"login_at" excel_name:"login_at"`
	LastLoginAt    float64 `json:"last_login_at" excel_name:"last_login_at"`
	CreatedAt      float64 `json:"created_at" excel_name:"created_at"`
	IsInitial      float64 `json:"is_initial" excel_name:"is_initial"`
	PaidAt         float64 `json:"paid_at" excel_name:"paid_at"`
	UpdateAt       float64 `json:"update_at" excel_name:"update_at"`
	LogoutAt       float64 `json:"logout_at" excel_name:"logout_at"`
}

func TestExcel(t *testing.T) {
	var stru []Resume
	info := Resume{
		Name:       "张三",
		Sex:        "男",
		Like:       "game",
		Age:        19,
		CreateTime: time.Now().Unix(),
		UpdateTime: time.Now(),
		Status:     2,
		Test:       "1",
		Test2:      "1",
		Test3:      "1",
		Test4:      "1",
		Test5:      "1",
		Test6:      "1",
		Test7:      "1",
		Test8:      "1",
		Test9:      "1",
		Test15:     "1",
		Test16:     "1",
		Test17:     "1",
		Test18:     "1",
		Test19:     "1",
	}

	for i := 0; i < 10; i++ {
		stru = append(stru, info)
	}
	e := Excel{}
	err := e.SaveExcel("Book1.xlsx", &stru)
	fmt.Println(err)
	return
}
func TestCsvReadExcel(t *testing.T) {
	file, err := os.Open("result.csv")
	if err != nil {
		fmt.Println(err)
	}
	defer file.Close()
	reader := csv.NewReader(file)
	for {
		record, err := reader.Read()
		if err == io.EOF {
			break
		} else if err != nil {
			fmt.Println("Error:", err)
			return
		}
		fmt.Println(record) // record has the type []string
	}

	return
}
func TestExcel_ReadExcel(t *testing.T) {
	e := Excel{
		Sheet: "Sheet1",
	}
	var stru []*Resume
	err := e.ReadExcel("Book1.xlsx", &stru)
	fmt.Println(err)
	for _, v := range stru {
		fmt.Println(v)
	}
	jsonStr, _ := json.Marshal(stru)
	fmt.Println(string(jsonStr))
	return
}

type chartDataRow struct {
	X     string  `json:"x"`
	Y     string  `json:"y"`
	Value float64 `json:"value"`
}
type chartData struct {
	Columns []string       `json:"columns"`
	Rows    []chartDataRow `json:"rows"`
}

func TestExcel_ReadCsv(t *testing.T) {
	e := Excel{}
	var stru []*HeatMapData
	err := e.ReadCsv("result.csv", &stru)
	fmt.Println(err)
	jsonStr, _ := json.Marshal(stru)
	fmt.Println(string(jsonStr))
	ChartData := chartData{}
	ChartData.Columns = []string{"x", "y", "value"}
	for _, v := range stru {
		info := &chartDataRow{}
		vStr, _ := json.Marshal(v)
		vMap := make(map[string]interface{})
		_ = json.Unmarshal(vStr, &vMap)
		info.X = v.Index
		for k, e := range vMap {
			if k == "index" {
				continue
			}
			info.Y = k
			info.Value = e.(float64)
			ChartData.Rows = append(ChartData.Rows, *info)
		}
	}
	jsonStr, _ = json.Marshal(ChartData)
	fmt.Println(string(jsonStr))
	return
}

func TestRunPythonScript(t *testing.T) {
	out, err := exec.Command("python3", "heatmap.py").Output()
	if err != nil {
		fmt.Println(err)
		return
	}
	result := string(out)
	fmt.Println("result:", result)
	if strings.Index(result, "success") == -1 {
		err = errors.New(fmt.Sprintf("error：%s", result))
	}
	fmt.Println("err:", err)
	return
}
func TestAddStr(t *testing.T) {

	fmt.Println(addStr("BZ", 3))
}

func Test360ReadExcel(t *testing.T) {
	f, err := excelize.OpenFile("result.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	// Get value from cell by given worksheet name and axis.
	cell, err := f.GetCellValue("result", "B2")
	if err != nil {
		fmt.Println(err)
		return
	}
	fmt.Println(cell)
	// Get all the rows in the Sheet1.
	rows, err := f.GetRows("result")
	for _, row := range rows {
		for _, colCell := range row {
			fmt.Print(colCell, "\t")
		}
		fmt.Println()
	}
}
func TestWriteCsv(t *testing.T) {
	var stru []Resume
	info := Resume{
		Name:       "张三",
		Sex:        "男",
		Like:       "game",
		Age:        19,
		CreateTime: time.Now().Unix(),
		UpdateTime: time.Now(),
		Status:     2,
		Test:       "1",
		Test2:      "1",
		Test3:      "1",
		Test4:      "1",
		Test5:      "1",
		Test6:      "1",
		Test7:      "1",
		Test8:      "1",
		Test9:      "1",
		Test15:     "1",
		Test16:     "1",
		Test17:     "1",
		Test18:     "1",
		Test19:     "1",
	}

	for i := 0; i < 10; i++ {
		stru = append(stru, info)
	}
	e := Excel{}
	err := e.SaveCsv("test_02.csv", &stru)
	fmt.Println(err)
	return
}
