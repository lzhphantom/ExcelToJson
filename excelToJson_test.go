/*
@Time : 2020/5/25 11:25
@Author : lzhphantom
@File : excelToJson_test
@Software: GoLand
*/
package ExcelToJson_test

import (
	"github.com/lzhphantom/ExcelToJson"
	"testing"
)

func TestAllSheetToJson(t *testing.T) {
	tests := []struct {
		filePath string
	}{
		{
			"./test.xlsx",
		},
	}

	for _, test := range tests {
		result, err := ExcelToJson.AllSheetToJson(test.filePath)
		if err != nil {
			t.Error("to json err", err)
		}
		t.Log(string(result))
	}
}

func TestASheetToJson(t *testing.T) {
	tests := []struct {
		filePath  string
		sheetName string
	}{
		{
			"./test.xlsx",
			"Sheet1",
		},
		{
			"./test.xlsx",
			"Sheet2",
		},
	}

	for _, test := range tests {
		result, err := ExcelToJson.ASheetToJson(test.filePath, test.sheetName)
		if err != nil {
			t.Error("to json err", err)
		}
		t.Log(string(result))
	}
}
