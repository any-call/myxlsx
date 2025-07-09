package myxlsx

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"reflect"
	"sort"
	"strconv"
	"strings"
	"time"
)

type (
	headerInfo struct {
		Path     string
		Title    string
		Order    int
		MaxWidth float64
	}
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

func ExportXLSXWithOptions[T any](list []T, filepath, sheetName string,
	getHeader func(fieldPath string) (string, int),
	getCellValue func(item T, fieldPath string) string) error {
	if len(list) == 0 {
		return fmt.Errorf("data is empty")
	}

	var sample T
	fieldPaths := extractFields(reflect.TypeOf(sample), "")

	var headers []headerInfo
	for _, path := range fieldPaths {
		title := path
		order := 9999

		if getHeader != nil {
			t, o := getHeader(path)
			if t == "" {
				continue
			}
			title = t
			order = o
		}

		headers = append(headers, headerInfo{
			Path:     path,
			Title:    title,
			Order:    order,
			MaxWidth: float64(len(title)) + 2, // 初始列宽
		})
	}

	sort.Slice(headers, func(i, j int) bool {
		return headers[i].Order < headers[j].Order
	})

	f := excelize.NewFile()
	sheetIdx, err := f.NewSheet(sheetName)
	if err != nil {
		return err
	}
	f.SetActiveSheet(sheetIdx)

	// 写 header
	for i, h := range headers {
		cell, _ := excelize.CoordinatesToCellName(i+1, 1)
		_ = f.SetCellValue(sheetName, cell, h.Title)
	}

	// 写内容
	for rowIdx, item := range list {
		v := reflect.ValueOf(item)

		for colIdx, h := range headers {
			val := ""

			if getCellValue != nil {
				val = getCellValue(item, h.Path)
			}

			if val == "" {
				raw := getValueByPath(v, h.Path)
				val = fmt.Sprintf("%v", raw)
			}

			// 更新最大列宽
			l := len(val)
			if float64(l) > headers[colIdx].MaxWidth {
				headers[colIdx].MaxWidth = float64(l) + 2
			}

			cell, _ := excelize.CoordinatesToCellName(colIdx+1, rowIdx+2)
			_ = f.SetCellValue(sheetName, cell, val)

			// 自动换行设置
			if strings.Contains(val, "\n") || strings.Contains(val, "\r") {
				style, _ := f.NewStyle(&excelize.Style{
					Alignment: &excelize.Alignment{WrapText: true},
				})
				_ = f.SetCellStyle(sheetName, cell, cell, style)
			}
		}
	}

	// 设置列宽
	for i, h := range headers {
		col, _ := excelize.ColumnNumberToName(i + 1)
		w := h.MaxWidth
		if w > 80 {
			w = 80
		}
		_ = f.SetColWidth(sheetName, col, col, w)
	}

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
