package myxlsx

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"reflect"
	"strconv"
	"time"
)

// ExportToXLSX 导出结构体切片为 XLSX 文件
func ExportToXLSX[T any](list []T, sheetName string, filepath string) error {
	if len(list) == 0 {
		return fmt.Errorf("empty data list")
	}

	f := excelize.NewFile()
	index, err := f.NewSheet(sheetName)
	if err != nil {
		return err
	}

	typeT := reflect.TypeOf(list[0])
	var headers []string
	var fieldIndex []int

	// 获取标签作为表头
	for i := 0; i < typeT.NumField(); i++ {
		field := typeT.Field(i)
		name := field.Tag.Get("xlsx")
		if name == "-" {
			continue
		}
		if name == "" {
			name = field.Name
		}
		headers = append(headers, name)
		fieldIndex = append(fieldIndex, i)
	}

	// 写入表头
	for col, name := range headers {
		cell, _ := excelize.CoordinatesToCellName(col+1, 1)
		f.SetCellValue(sheetName, cell, name)
	}

	// 写入数据
	for row, item := range list {
		v := reflect.ValueOf(item)
		for col, i := range fieldIndex {
			val := v.Field(i).Interface()
			cell, _ := excelize.CoordinatesToCellName(col+1, row+2)
			f.SetCellValue(sheetName, cell, val)
		}
	}

	f.SetActiveSheet(index)
	return f.SaveAs(filepath)
}

// ImportFromXLSX 从 XLSX 导入为结构体切片
func ImportFromXLSX[T any](filepath string, sheetName string) ([]T, error) {
	f, err := excelize.OpenFile(filepath)
	if err != nil {
		return nil, err
	}

	rows, err := f.GetRows(sheetName)
	if err != nil || len(rows) < 2 {
		return nil, fmt.Errorf("invalid or empty sheet")
	}

	headers := rows[0]
	var result []T
	typeT := reflect.TypeOf((*T)(nil)).Elem()

	fieldMap := map[string]int{}
	for i := 0; i < typeT.NumField(); i++ {
		field := typeT.Field(i)
		name := field.Tag.Get("xlsx")
		if name == "-" {
			continue
		}
		if name == "" {
			name = field.Name
		}
		fieldMap[name] = i
	}

	// 从第二行开始读取
	for _, row := range rows[1:] {
		v := reflect.New(typeT).Elem()
		for colIdx, colVal := range row {
			if colIdx >= len(headers) {
				continue
			}
			fieldName := headers[colIdx]
			if fieldIdx, ok := fieldMap[fieldName]; ok {
				field := v.Field(fieldIdx)
				if field.CanSet() {
					switch field.Kind() {
					case reflect.String:
						field.SetString(colVal)
						break
					case reflect.Int, reflect.Int64:
						i, _ := strconv.ParseInt(colVal, 10, 64)
						field.SetInt(i)
						break
					case reflect.Float64:
						fval, _ := strconv.ParseFloat(colVal, 64)
						field.SetFloat(fval)
						break
					case reflect.Bool:
						b, _ := strconv.ParseBool(colVal)
						field.SetBool(b)
						break
					case reflect.Struct:
						if field.Type() == reflect.TypeOf(time.Time{}) {
							if t, err := time.Parse("2006-01-02 15:04:05", colVal); err == nil {
								field.Set(reflect.ValueOf(t))
							} else if t, err := time.Parse("2006-01-02", colVal); err == nil {
								field.Set(reflect.ValueOf(t))
							}
						}
						break

					}
				}
			}
		}
		result = append(result, v.Interface().(T))
	}

	return result, nil
}
