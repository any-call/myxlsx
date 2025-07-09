package myxlsx

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"reflect"
	"strconv"
	"strings"
	"time"
)

// ExportToXLSX 导出结构体切片为 XLSX 文件
func ExportToXLSX[T any](list []T, filepath, sheetName string) error {
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

// ExportToXLSX 导出结构体切片为 XLSX 文件
func ExportToXLSXEx[T any](list []T, filepath, sheetName string,
	getHeader func(fieldPath string) string,
	getCellValue func(item T, fieldPath string) string) error {
	if len(list) == 0 {
		return fmt.Errorf("empty data list")
	}

	// 反射提取字段路径
	var sample T
	fieldPaths := extractFields(reflect.TypeOf(sample), "")

	// 确定导出的字段（根据 getHeader）
	var headers []string
	var exportFields []string

	for _, path := range fieldPaths {
		if getHeader != nil {
			title := getHeader(path)
			if title == "" {
				continue // 忽略字段
			}
			headers = append(headers, title)
			exportFields = append(exportFields, path)
		} else {
			headers = append(headers, path)
			exportFields = append(exportFields, path)
		}
	}

	f := excelize.NewFile()
	index, err := f.NewSheet(sheetName)
	if err != nil {
		return err
	}
	f.SetActiveSheet(index)

	// 写入 header
	for i, title := range headers {
		cell, _ := excelize.CoordinatesToCellName(i+1, 1)
		f.SetCellValue(sheetName, cell, title)
	}

	// 写入数据行
	for rowIndex, item := range list {
		v := reflect.ValueOf(item)

		for colIndex, fieldPath := range exportFields {
			var val string
			// 优先走回调
			if getCellValue != nil {
				val = getCellValue(item, fieldPath)
			} else {
				// 默认反射值
				raw := getValueByPath(v, fieldPath)
				val = fmt.Sprintf("%v", raw)
			}

			cell, _ := excelize.CoordinatesToCellName(colIndex+1, rowIndex+2)
			f.SetCellValue(sheetName, cell, val)
		}
	}
	// 保存文件
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

// 递归展开字段路径
func extractFields(t reflect.Type, prefix string) []string {
	var fields []string

	for i := 0; i < t.NumField(); i++ {
		field := t.Field(i)

		// 忽略未导出字段
		if !field.IsExported() {
			continue
		}

		// 当前字段路径
		fieldPath := field.Name
		if prefix != "" {
			fieldPath = prefix + "." + field.Name
		}

		ft := field.Type
		if ft.Kind() == reflect.Struct && ft.PkgPath() != "" {
			// 是用户自定义结构体，递归展开
			subFields := extractFields(ft, fieldPath)
			fields = append(fields, subFields...)
		} else {
			fields = append(fields, fieldPath)
		}
	}

	return fields
}

// 根据字段路径获取结构体中的值（支持嵌套）
func getValueByPath(v reflect.Value, fieldPath string) any {
	fields := strings.Split(fieldPath, ".")

	for _, name := range fields {
		if v.Kind() == reflect.Pointer {
			v = v.Elem()
		}
		if v.Kind() != reflect.Struct {
			return nil
		}
		v = v.FieldByName(name)
	}

	if v.IsValid() && v.CanInterface() {
		return v.Interface()
	}
	return nil
}
