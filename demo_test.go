package myxlsx

import (
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

	if err := ExportToXLSX(data, "Sheet1", "/Users/luisjin/Desktop/用户表.xlsx"); err != nil {
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
