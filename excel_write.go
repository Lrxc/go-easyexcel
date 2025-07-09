package easyexcel

import (
	"bytes"
	"fmt"
	"github.com/xuri/excelize/v2"
	"reflect"
	"strings"
)

// 将结构体数组导出为Excel文件
//
// @data: 结构体数组
// @sheetName: 工作表名称
// @Return: Excel文件字节数组, 错误信息
func ExcelWrite(data interface{}) ([]byte, error) {
	return ExcelWriteWithSheetName(data, "Sheet1")
}

func ExcelWriteWithSheetName(data interface{}, sheetName string) ([]byte, error) {
	f := excelize.NewFile()
	defer f.Close()

	if sheetName != "" {
		f.SetSheetName("Sheet1", sheetName) //默认有个sheet1,直接重命名,避免残留默认的sheet1
	}
	// 创建工作表
	//index, err := f.NewSheet(sheetName)
	//if err != nil {
	//	return nil, err
	//}
	//f.SetActiveSheet(index)

	// 通过反射获取结构体信息和数据
	sliceValue := reflect.ValueOf(data)
	if sliceValue.Kind() != reflect.Slice {
		return nil, fmt.Errorf("data must be a slice")
	}

	// 处理空数组情况
	if sliceValue.Len() == 0 {
		return nil, fmt.Errorf("empty data slice")
	}

	// 获取第一个元素的结构体类型
	elemType := sliceValue.Index(0).Type()
	if elemType.Kind() == reflect.Ptr {
		elemType = elemType.Elem()
	}

	// 写入表头（跳过忽略字段）
	colIndex := 0
	for i := 0; i < elemType.NumField(); i++ {
		field := elemType.Field(i)
		//是否忽略字段
		if shouldIgnoreField(field) {
			continue
		}
		//获取表头名称
		header := getHeaderName(field)

		cell, _ := excelize.CoordinatesToCellName(colIndex+1, 1)
		f.SetCellValue(sheetName, cell, header)
		f.SetCellStyle(sheetName, cell, cell, headStyle(f))
		colIndex++
	}

	// 写入数据行
	for row := 0; row < sliceValue.Len(); row++ {
		elem := sliceValue.Index(row)
		if elem.Kind() == reflect.Ptr {
			elem = elem.Elem()
		}

		colIndex := 0
		for col := 0; col < elem.NumField(); col++ {
			field := elemType.Field(col)
			//是否忽略字段
			if shouldIgnoreField(field) {
				continue
			}

			cell, _ := excelize.CoordinatesToCellName(colIndex+1, row+2)
			fieldValue := elem.Field(col).Interface()
			f.SetCellValue(sheetName, cell, transValue(field, fieldValue))
			colIndex++
		}
	}

	// 将Excel文件写入缓冲区
	var buf bytes.Buffer
	if err := f.Write(&buf); err != nil {
		return nil, err
	}

	return buf.Bytes(), nil
}

// 表头样式
func headStyle(f *excelize.File) int {
	headStyle, _ := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Bold: true,
			Size: 12,
		},
	})
	return headStyle
}

// 设置单元格的值（处理特殊类型）
func transValue(field reflect.StructField, fieldValue any) interface{} {
	tag := field.Tag.Get(TAG_EASYEXCEL_NAME)
	split := strings.Split(tag, ",")

	convertTag := ""
	for _, part := range split {
		if strings.HasPrefix(part, TAG_EASYEXCEL_CONVERT) {
			convertTag = strings.TrimPrefix(part, TAG_EASYEXCEL_CONVERT)
			break
		}
	}

	// 检查是否有转换器
	if convertTag != "" {
		// 反射调用转换函数
		dynamic, err := TransConvert(convertTag+EASYEXCEL_CONVERT_WRITE, fieldValue)
		if err != nil {
			dynamic, err = TransConvert(convertTag, fieldValue)
			if err != nil {
				fmt.Printf("excel convent err: field=%s, %s", field.Name, err)
				return fieldValue
			}
		}
		return dynamic[0]
	}
	return fieldValue
}

// 判断是否应该忽略字段
func shouldIgnoreField(field reflect.StructField) bool {
	tag := field.Tag.Get(TAG_EASYEXCEL_NAME)
	return tag == "" || tag == "-"
}

// 获取表头名称(excel:"状态,convert=UserConv.Status")
func getHeaderName(field reflect.StructField) string {
	tag := field.Tag.Get(TAG_EASYEXCEL_NAME)
	if tag != "" && tag != "-" {
		split := strings.Split(tag, ",")
		return split[0]
	}
	return field.Name
}
