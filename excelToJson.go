/*
@Time : 2020/5/25 10:54
@Author : lzhphantom
@File : excelToJson
@Software: GoLand
*/
package ExcelToJson

import (
	"encoding/json"
	"github.com/tealeg/xlsx"
)

type sheetJSONMap []map[string]interface{}

// exchange all excel sheet information to json
// with sheetName
func AllSheetToJson(xlsxPath string) ([]byte, error) {
	xlsxFile, err := xlsx.OpenFile(xlsxPath)
	if err != nil {
		return nil, err
	}
	infos := make(map[string]sheetJSONMap)
	for _, sheet := range xlsxFile.Sheets {
		//fmt.Println("sheet name:", sheet.Name)
		info := make(sheetJSONMap, 0)
		markMap := make(map[int]string)
		for rowIndex, row := range sheet.Rows {
			var jsonMap map[string]interface{}
			if rowIndex != 0 {
				jsonMap = make(map[string]interface{})
			}
			for cellIndex, cell := range row.Cells {
				if rowIndex == 0 {
					markMap[cellIndex] = cell.Value
				} else {
					if jsonMap != nil {
						jsonMap[markMap[cellIndex]] = cell.Value
					}
				}
			}
			if jsonMap != nil {
				info = append(info, jsonMap)
			}
		}
		infos[sheet.Name] = info
	}
	jsonBytes, err := json.MarshalIndent(&infos, "", "	")
	if err != nil {
		return nil, err
	}

	return jsonBytes, nil
}

// exchange a excel sheet information to json
// by sheetName
func ASheetToJson(xlsxPath string, sheetName string) ([]byte, error) {
	xlsxFile, err := xlsx.OpenFile(xlsxPath)
	if err != nil {
		return nil, err
	}

	sheet := xlsxFile.Sheet[sheetName]
	info := make(sheetJSONMap, 0)
	markMap := make(map[int]string)
	for rowIndex, row := range sheet.Rows {
		var jsonMap map[string]interface{}
		if rowIndex != 0 {
			jsonMap = make(map[string]interface{})
		}
		for cellIndex, cell := range row.Cells {
			if rowIndex == 0 {
				markMap[cellIndex] = cell.Value
			} else {
				if jsonMap != nil {
					jsonMap[markMap[cellIndex]] = cell.Value
				}
			}
		}
		if jsonMap != nil {
			info = append(info, jsonMap)
		}
	}
	jsonBytes, err := json.MarshalIndent(&info, "", "	")
	if err != nil {
		return nil, err
	}

	return jsonBytes, nil
}
