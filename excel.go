package excel

import (
	"errors"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"reflect"
	"strconv"
	"strings"
	"time"
)

type Excel struct {
	f *excelize.File
	Sheet string
}
func (e *Excel) ReadExcel(file string,data interface{})(err error){
	if e.Sheet == "" {
		e.Sheet = "Sheet1"
	}
	if reflect.ValueOf(data).Kind() != reflect.Ptr || reflect.TypeOf(data).Elem().Kind() != reflect.Slice {
		err = errors.New("参数错误")
		return err
	}
	e.f ,err = excelize.OpenFile(file)
	if err != nil {
		return err
	}
	v := reflect.ValueOf(data).Elem()
	t := reflect.TypeOf(data).Elem().Elem()
	if t.Kind() == reflect.Ptr {
		t = t.Elem()
	}
	// Get all the rows in the Sheet1.
	rows, err := e.f.GetRows(e.Sheet)
	var map_data = make(map[int]int)
	var map_arr = make(map[int]map[string]int)
	var v_index int = 0
	for x, row := range rows {
		if x == 0 {

			for y, colCell := range row {
				if x == 0 {
					null_num := 0
					for i := 0; i < t.NumField(); i++ {
						excelName := t.Field(i).Tag.Get("excelName")
						if excelName == "" {
							null_num++
							continue
						}
						if excelName == colCell {
							map_data[i] = y
						}
						if enums_str := t.Field(i).Tag.Get("enums") ; enums_str != "" {
							var map_v = make(map[string]int)
							enums_arr := strings.Split(enums_str, ",")
							for _, v := range enums_arr {
								enums_v := strings.Split(v, ":")
								key, _ := strconv.Atoi(enums_v[0])
								map_v[enums_v[1]] = key
							}
							map_arr[i] = map_v
						}
					}
				}

			}
			continue
		}
		var subv reflect.Value
		subv = reflect.New(t)
		if v.Type().Elem().Kind() != reflect.Ptr {
			subv = subv.Elem()
		}
		v2 := reflect.Append(v,subv)
		v.Set(v2)
		v_index_value := v.Index(v_index)
		if v_index_value.Type().Kind() == reflect.Ptr {
			v_index_value = v_index_value.Elem()
		}
		for i := 0; i < v_index_value.NumField();i++ {
			if _, ok := map_data[i]; !ok {
				continue
			}
			excel_value := row[map_data[i]]
			switch v_index_value.Field(i).Type().Kind() {
			case reflect.String:
				v_index_value.Field(i).Set(reflect.ValueOf(excel_value))
				break
			case reflect.Int:
				data_v ,_ := strconv.Atoi(excel_value)
				v_index_value.Field(i).Set(reflect.ValueOf(data_v))
				break
			case reflect.Int8:
				data_v ,_ := strconv.Atoi(excel_value)
				if enums_str := t.Field(i).Tag.Get("enums") ; enums_str != "" {
					map_v := map_arr[i]
					if _ ,ok := map_v[excel_value] ; ok{
						data_v = map_v[excel_value]
					}
				}
				v_index_value.Field(i).Set(reflect.ValueOf(int8(data_v)))
				break
			case reflect.Int16:
				data_v ,_ := strconv.Atoi(excel_value)
				v_index_value.Field(i).Set(reflect.ValueOf(int16(data_v)))
				break
			case reflect.Int32:
				data_v ,_ := strconv.Atoi(excel_value)
				v_index_value.Field(i).Set(reflect.ValueOf(int32(data_v)))
				break
			case reflect.Int64:
				data_v ,_ := strconv.Atoi(excel_value)
				if t.Field(i).Tag.Get("excelTime") == "int" || t.Field(i).Tag.Get("excelTime") == "int64" {
					local, _ := time.LoadLocation("Local")
					t, _ := time.ParseInLocation("2006-01-02 15:04:05", excel_value, local)
					data_v = int(t.Unix())
				}
				v_index_value.Field(i).Set(reflect.ValueOf(int64(data_v)))
				break
			case reflect.Uint:
				data_v ,_ := strconv.Atoi(excel_value)
				v_index_value.Field(i).Set(reflect.ValueOf(uint(data_v)))
				break
			case reflect.Uint8:
				data_v ,_ := strconv.Atoi(excel_value)
				v_index_value.Field(i).Set(reflect.ValueOf(uint8(data_v)))
				break
			case reflect.Uint16:
				data_v ,_ := strconv.Atoi(excel_value)
				v_index_value.Field(i).Set(reflect.ValueOf(uint16(data_v)))
				break
			case reflect.Uint32:
				data_v ,_ := strconv.Atoi(excel_value)
				v_index_value.Field(i).Set(reflect.ValueOf(uint32(data_v)))
				break
			case reflect.Uint64:
				data_v ,_ := strconv.Atoi(excel_value)
				v_index_value.Field(i).Set(reflect.ValueOf(uint64(data_v)))
				break
			case reflect.Float32:
				data_v,_ := strconv.ParseFloat(excel_value,64)
				v_index_value.Field(i).Set(reflect.ValueOf(float32(data_v)))
				break
			case reflect.Float64:
				data_v,_ := strconv.ParseFloat(excel_value,64)
				v_index_value.Field(i).Set(reflect.ValueOf(data_v))
				break
			case reflect.Bool:
				data_v,_ := strconv.ParseBool(excel_value)
				v_index_value.Field(i).Set(reflect.ValueOf(data_v))
				break
			case reflect.TypeOf(time.Time{}).Kind():
				local, _ := time.LoadLocation("Local")
				t, _ := time.ParseInLocation("2006-01-02 15:04:05", excel_value, local)
				v_index_value.Field(i).Set(reflect.ValueOf(t))
				break
			default:
				//v_index_value.Field(i).Set(reflect.ValueOf(0))
				//break
			}
		}
		v_index++

	}
	return
}
func (e *Excel) SaveExcel(file string,data interface{})(err error){
	if e.Sheet == "" {
		e.Sheet = "Sheet1"
	}
	if reflect.ValueOf(data).Kind() != reflect.Ptr || reflect.TypeOf(data).Elem().Kind() != reflect.Slice {
		err = errors.New("参数错误")
		return err
	}
	e.f = excelize.NewFile()
	index := e.f.NewSheet(e.Sheet)
	t := reflect.TypeOf(data).Elem().Elem()
	if t.Kind() == reflect.Ptr {
		t = t.Elem()
	}
	null_num := 0
	var map_arr = make(map[int]map[int]string)
	for i := 0; i < t.NumField(); i++ {
		excelName := t.Field(i).Tag.Get("excelName")
		if excelName == "" {
			null_num++
			continue
		}
		enums_str := t.Field(i).Tag.Get("enums")
		if enums_str != "" {
			map_data := make(map[int]string)
			enums_arr := strings.Split(enums_str,",")
			for _,v := range enums_arr {
				enums_v := strings.Split(v,":")
				key ,_ := strconv.Atoi(enums_v[0])
				map_data[key] = enums_v[1]
			}
			map_arr[i] = map_data
		}
		axis := fmt.Sprintf("%s1",addStr("A",int32(i-null_num)))
		e.f.SetCellValue(e.Sheet,axis,excelName)
	}
	num := reflect.ValueOf(data).Elem().Len()
	if num > 0 {
		 for i := 0; i < num; i++ {
		 	value := reflect.ValueOf(data).Elem().Index(i)
		 	if reflect.ValueOf(data).Elem().Index(i).Type().Kind() == reflect.Ptr {
				value = value.Elem()
			}
			t := value.Type()
			null_num := 0
			for x := 0; x < t.NumField(); x++ {
				excelName := t.Field(x).Tag.Get("excelName")
				if excelName == "" {
					null_num++
					continue
				}
				cell_value := value.Field(x)
				enums_str := t.Field(x).Tag.Get("enums")
				if enums_str != "" {
					map_data := map_arr[x]
					if _ ,ok := map_data[int(value.Field(x).Int())] ; ok{
						cell_value = reflect.ValueOf(map_data[int(value.Field(x).Int())])
					}
				}
				excelTime := t.Field(x).Tag.Get("excelTime")
				switch excelTime {
				case "time":
					val_str := fmt.Sprintf("%v",value.Field(x))
					cell_value = reflect.ValueOf(val_str[:19])
					break
				case "int","int64":
					t := time.Unix(value.Field(x).Int(), 0).Format("2006-01-02 15:04:05")
					cell_value = reflect.ValueOf(t)
					break
				}
				axis := fmt.Sprintf("%s%d",addStr("A",int32(x-null_num)),i + 2)
				e.f.SetCellValue(e.Sheet,axis,cell_value)
			}
		 }
	}
	// Set value of a cell.
	// Set active sheet of the workbook.
	e.f.SetActiveSheet(index)
	// Save spreadsheet by the given path.
	err = e.f.SaveAs(file)
	return
}
func addStr(v string,add int32) string{
	x := []rune(v)
	for index := range x {
		x[index] = x[index] + add
	}
	return string(x)
}
