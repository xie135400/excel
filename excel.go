package excel

import (
	"errors"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"reflect"
	"strconv"
)

type Excel struct {
	f *excelize.File
	Sheet1 string
}
func (e *Excel) ReadExcel(file string,data interface{})(err error){
	if e.Sheet1 == "" {
		e.Sheet1 = "Sheet1"
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
	c := reflect.TypeOf(data).Elem().Elem()
	t := reflect.TypeOf(data).Elem().Elem()
	// Get all the rows in the Sheet1.
	rows, err := e.f.GetRows("Sheet1")
	var map_data = make(map[int]int)
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
					}
				}

			}
			continue
		}
		var subv reflect.Value
		subv = reflect.New(c).Elem()
		v2 := reflect.Append(v,subv)
		v.Set(v2)
		for i := 0; i < v.Index(v_index).NumField();i++ {
			if _, ok := map_data[i]; !ok {
				continue
			}
			excel_value := row[map_data[i]]
			switch v.Index(v_index).Field(i).Type().Kind() {
			case reflect.String:
				v.Index(v_index).Field(i).Set(reflect.ValueOf(excel_value))
				break
			case reflect.Int:
				data_v ,_ := strconv.Atoi(excel_value)
				v.Index(v_index).Field(i).Set(reflect.ValueOf(data_v))
				break
			case reflect.Int8:
				data_v ,_ := strconv.Atoi(excel_value)
				v.Index(v_index).Field(i).Set(reflect.ValueOf(int8(data_v)))
				break
			case reflect.Int16:
				data_v ,_ := strconv.Atoi(excel_value)
				v.Index(v_index).Field(i).Set(reflect.ValueOf(int16(data_v)))
				break
			case reflect.Int32:
				data_v ,_ := strconv.Atoi(excel_value)
				v.Index(v_index).Field(i).Set(reflect.ValueOf(int32(data_v)))
				break
			case reflect.Int64:
				data_v ,_ := strconv.Atoi(excel_value)
				v.Index(v_index).Field(i).Set(reflect.ValueOf(int64(data_v)))
				break
			case reflect.Uint:
				data_v ,_ := strconv.Atoi(excel_value)
				v.Index(v_index).Field(i).Set(reflect.ValueOf(uint(data_v)))
				break
			case reflect.Uint8:
				data_v ,_ := strconv.Atoi(excel_value)
				v.Index(v_index).Field(i).Set(reflect.ValueOf(uint8(data_v)))
				break
			case reflect.Uint16:
				data_v ,_ := strconv.Atoi(excel_value)
				v.Index(v_index).Field(i).Set(reflect.ValueOf(uint16(data_v)))
				break
			case reflect.Uint32:
				data_v ,_ := strconv.Atoi(excel_value)
				v.Index(v_index).Field(i).Set(reflect.ValueOf(uint32(data_v)))
				break
			case reflect.Uint64:
				data_v ,_ := strconv.Atoi(excel_value)
				v.Index(v_index).Field(i).Set(reflect.ValueOf(uint64(data_v)))
				break
			case reflect.Float32:
				data_v,_ := strconv.ParseFloat(excel_value,64)
				v.Index(v_index).Field(i).Set(reflect.ValueOf(float32(data_v)))
				break
			case reflect.Float64:
				data_v,_ := strconv.ParseFloat(excel_value,64)
				v.Index(v_index).Field(i).Set(reflect.ValueOf(data_v))
				break
			case reflect.Bool:
				data_v,_ := strconv.ParseBool(excel_value)
				v.Index(v_index).Field(i).Set(reflect.ValueOf(data_v))
				break
			default:
				err = errors.New("不支持的数据类型")
				return err
			}
		}
		v_index++

	}
	return
}
func (e *Excel) SaveExcel(file string,data interface{})(err error){
	if e.Sheet1 == "" {
		e.Sheet1 = "Sheet1"
	}
	if reflect.ValueOf(data).Kind() != reflect.Ptr || reflect.TypeOf(data).Elem().Kind() != reflect.Slice {
		err = errors.New("参数错误")
		return err
	}
	e.f = excelize.NewFile()
	index := e.f.NewSheet(e.Sheet1)
	t := reflect.TypeOf(data).Elem().Elem()
	null_num := 0
	for i := 0; i < t.NumField(); i++ {
		excelName := t.Field(i).Tag.Get("excelName")
		if excelName == "" {
			null_num++
			continue
		}
		axis := fmt.Sprintf("%s1",addStr("A",int32(i-null_num)))
		e.f.SetCellValue(e.Sheet1,axis,excelName)
	}
	num := reflect.ValueOf(data).Elem().Len()
	if num > 0 {
		 for i := 0; i < num; i++ {
		 	value := reflect.ValueOf(data).Elem().Index(i)
			t := value.Type()
			null_num := 0
			for x := 0; x < t.NumField(); x++ {
				excelName := t.Field(x).Tag.Get("excelName")
				if excelName == "" {
					null_num++
					continue
				}
				axis := fmt.Sprintf("%s%d",addStr("A",int32(x-null_num)),i + 2)
				e.f.SetCellValue(e.Sheet1,axis,value.Field(x))
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
