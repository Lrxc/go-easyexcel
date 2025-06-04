package easyexcel

import (
	"log"
	"os"
	"slices"
	"testing"
)

type Database struct {
	DbUrl  string `excel:"数据库地址"`                            //excel字段名
	DbUser string `excel:"数据库用户名"`                           //excel字段名
	DbPwd  string `excel:"-"`                                //忽略字段
	Status int8   `excel:"状态" convert:"DatabaseConv.Status"` //设置转换器
}

type DatabaseConv struct {
}

// excel写入转换(添加_Write)
func (DatabaseConv) Status_Write(value any) any {
	var arr = []string{"异常", "正常"}
	return arr[value.(int8)]
}

// excel读取转换(添加_Read)
func (DatabaseConv) Status_Read(value any) any {
	var arr = []string{"异常", "正常"}
	return slices.Index(arr, value.(string))
}

// 生成excel
func TestGen(t *testing.T) {
	d := Database{
		DbUrl:  "127.0.0.1:3306",
		DbUser: "user",
		DbPwd:  "user",
		Status: 1,
	}

	RegConvert(DatabaseConv{})

	gen, err := ExcelWrite([]Database{d})
	if err != nil {
		log.Println("err: ", err)
	}
	os.WriteFile("test.xlsx", gen, os.ModePerm)
}

// 解析excel
func TestParse(t *testing.T) {
	RegConvert(DatabaseConv{})

	var dbs []Database
	err := ExcelRead("test.xlsx", &dbs)
	if err != nil {
		log.Println("err: ", err)
	}

	for _, db := range dbs {
		log.Println(db)
	}
}
