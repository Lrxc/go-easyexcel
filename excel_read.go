package easyexcel

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"reflect"
	"strconv"
)

func ExcelRead(filePath string, dest interface{}) error {
	return ExcelReadWithSheetName(filePath, "", dest)
}

// 将excel 解析为 结构体数组
func ExcelReadWithSheetName(filePath string, sheetName string, dest interface{}) error {
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return fmt.Errorf("excel open err: %v", err)
	}
	defer f.Close()

	destValue := reflect.ValueOf(dest)
	if destValue.Kind() != reflect.Ptr || destValue.Elem().Kind() != reflect.Slice {
		return fmt.Errorf("dest must is ptr")
	}

	sliceType := destValue.Elem().Type().Elem()
	if sliceType.Kind() != reflect.Struct {
		return fmt.Errorf("dest must is struct")
	}

	if sheetName == "" {
		//获取所有的sheet名
		sheetName = f.GetSheetList()[0]
	}
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return fmt.Errorf("sheet name err: %v", err)
	}

	if len(rows) == 0 {
		return fmt.Errorf("excel rows can't zero")
	}

	// 构建字段映射：Excel列名 -> 结构体字段信息
	headers := rows[0]
	fieldInfos := make([]*fieldInfo, len(headers))

	for i, header := range headers {
		fieldInfos[i] = findFieldByHeader(sliceType, header)
	}

	// 处理数据行
	for rowIdx, row := range rows[1:] {
		newItem := reflect.New(sliceType).Elem()

		for colIdx, cellValue := range row {
			info := fieldInfos[colIdx]
			if info == nil || info.ignore {
				continue
			}

			field := newItem.Field(info.index)
			if !field.CanSet() {
				continue
			}

			var dynamic interface{} = cellValue
			if info.convert != "" {
				// 应用转换器
				res, _ := TransConvert(info.convert+EASYEXCEL_CONVERT_READ, cellValue)
				dynamic = res[0]
			}

			if err := setFieldValue(field, dynamic); err != nil {
				return fmt.Errorf("row:%d col:%d val:%s err:%v", rowIdx+2, colIdx+1, headers[colIdx], err)
			}
		}

		destValue.Elem().Set(reflect.Append(destValue.Elem(), newItem))
	}
	return nil
}

// 字段信息结构
type fieldInfo struct {
	index   int    // 字段索引
	ignore  bool   // 是否忽略
	convert string // 转换规则
}

// 根据Excel列名查找结构体字段
func findFieldByHeader(typ reflect.Type, header string) *fieldInfo {
	for i := 0; i < typ.NumField(); i++ {
		field := typ.Field(i)

		tag := field.Tag.Get(TAG_EASYEXCEL_NAME)
		convert := field.Tag.Get(TAG_EASYEXCEL_CONVERT)

		if tag == "" || tag != header {
			continue
		}

		return &fieldInfo{
			index:   i,
			ignore:  tag == "-",
			convert: convert,
		}
	}
	return nil
}

// 设置字段值
func setFieldValue(field reflect.Value, value interface{}) error {
	if value == nil {
		return nil
	}

	val := reflect.ValueOf(value)
	//自动转换类型
	if val.Type().ConvertibleTo(field.Type()) {
		field.Set(val.Convert(field.Type()))
		return nil
	}

	switch field.Kind() {
	case reflect.Int8:
		field.SetInt(int64(value.(int8)))
	case reflect.Int16:
		field.SetInt(int64(value.(int16)))
	case reflect.Int32:
		field.SetInt(int64(value.(int32)))
	case reflect.Int64:
		atoi, _ := strconv.Atoi(fmt.Sprintf("%s", value))
		field.SetInt(int64(atoi))
	case reflect.String:
		field.SetString(value.(string))
	}

	return nil
}
