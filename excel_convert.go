package easyexcel

import (
	"fmt"
	"reflect"
	"strings"
)

// 对象注册表
var funcRegistry = make(map[string]interface{})

// 所有的转换器都要实现该接口
type IBaseConvert interface {
	EasyExcelConvert()
}

// 注册转换器
func RegConvert(converts ...IBaseConvert) {
	for _, convert := range converts {
		beanName := reflect.TypeOf(convert).Name()
		funcRegistry[beanName] = convert
	}
}

// 反射执行转换器
func TransConvert(convertTag string, args ...interface{}) ([]interface{}, error) {
	split := strings.Split(convertTag, ".")
	beanName := split[0]
	methodName := split[1]

	// 查找对象
	fn, ok := funcRegistry[beanName]
	if !ok {
		return nil, fmt.Errorf("bean not fount: %s", convertTag)
	}

	// 准备反射参数
	var in []reflect.Value
	for _, arg := range args {
		in = append(in, reflect.ValueOf(arg))
	}

	// 调用函数
	method := reflect.ValueOf(fn).MethodByName(methodName)
	if !method.IsValid() {
		return nil, fmt.Errorf("method %s not found", methodName)
	}
	out := method.Call(in)

	// 转换返回值
	var results []interface{}
	for _, val := range out {
		results = append(results, val.Interface())
	}

	return results, nil
}
