package myxlsx

import (
	"fmt"
	"testing"
	"time"
)

func TestExportToXLSX(t *testing.T) {
	type User struct {
		ID    int
		Name  string
		Email string
		Time  time.Time
		Phone string `xlsx:"-"`
	}

	data := []User{
		{ID: 1, Name: "Alice", Email: "alice@example.com", Time: time.Now()},
		{ID: 2, Name: "Tom", Email: "tom@example.com", Time: time.Now()},
	}

	if err := ExportToXLSX(data, "/Users/luisjin/Desktop/用户表.xlsx", "Sheet1"); err != nil {
		t.Error(err)
		return
	}

	t.Log("export ok")
}

func TestExportToXLSX1(t *testing.T) {
	type MoneyRec struct {
		Id        int64 //是否可空:NO
		UserId    int64
		TxID      string
		AddTime   int64
		Symbol    string
		MoneyType int
		Money     float64
		Remark    string
	}
	type User struct {
		MoneyRec
		ID    int
		Name  string
		Email string
		Time  time.Time
		Phone string `xlsx:"-"`
	}

	data := []User{
		{ID: 1, Name: "Alice", Email: "alice@example.com", Time: time.Now()},
		{ID: 2, Name: "Tom", Email: "tom@example.com", Time: time.Now()},
	}

	if err := ExportXLSXWithOptions(data, "/Users/luisjin/Desktop/用户表1.xlsx",
		"Sheet1",
		func(fieldPath string) string {
			fmt.Println("fieldPath is :", fieldPath)
			switch fieldPath {
			case "MoneyRec.Id":
				return ""
			default:
				break
			}
			return fieldPath
		},
		func(item User, fieldPath string) string {
			fmt.Println("fieldPath is :", fieldPath, "item is :", item)
			return fieldPath
		},
	); err != nil {
		t.Error(err)
		return
	}

	t.Log("export ok")
}

func TestImportFromXLSX(t *testing.T) {
	type User struct {
		ID    int
		Name  string
		Email string
		Time  time.Time
		Phone string `xlsx:"-"`
	}
	list, err := ImportFromXLSX[User]("/Users/luisjin/Desktop/用户表.xlsx", "Sheet1")
	if err != nil {
		t.Error(err)
		return
	}

	t.Log("list is :", list)
}
