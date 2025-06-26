# Go EasyExcel

# 使用
```shell
go get github.com/Lrxc/go-easyexcel v1.0.10
```

```go
//写入excel
easyexcel.ExcelWrite()
//读取excel
easyexcel.ExcelRead()
```

# 示例
```go
package easyexcel

import (
	"log"
	"os"
	"slices"
	"testing"
)

type Database struct {
	DbUrl  string `excel:"数据库地址"`                          //excel字段名
	DbUser string `excel:"数据库用户名"`                         //excel字段名
	DbPwd  string `excel:"-"`                              //忽略字段
	Status int8   `excel:"状态,convert=DatabaseConv.Status"` //设置转换器
}

type DatabaseConv struct{}

func (DatabaseConv) EasyExcelConvert() {}

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

func Init() {
	//注册自定义转换器
	RegConvert(DatabaseConv{})
}

func TestExcel(t *testing.T) {
	TestGen(t)
	TestParse(t)
}

// 生成excel
func TestGen(t *testing.T) {
	Init()

	d := []Database{
		Database{
			DbUrl:  "127.0.0.1:3306",
			DbUser: "user",
			DbPwd:  "user",
			Status: 1,
		},
		Database{
			DbUrl:  "127.0.0.1:5432",
			DbUser: "dev",
			DbPwd:  "dev",
			Status: 0,
		},
	}

	gen, err := ExcelWrite(d)
	if err != nil {
		log.Println("err: ", err)
	}
	os.WriteFile("test.xlsx", gen, os.ModePerm)
}

// 解析excel
func TestParse(t *testing.T) {
	Init()

	var dbs []Database
	err := ExcelRead("test.xlsx", &dbs)
	if err != nil {
		log.Println("err: ", err)
	}

	for _, db := range dbs {
		log.Printf("%+v", db)
	}
}

```