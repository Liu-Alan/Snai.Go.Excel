package main

import (
	"fmt"
	"io/ioutil"
	"strings"

	"Snai.Go.Excel/entities"
	"Snai.Go.Excel/utils"
)

func main() {
	dirPath := "excel"
	files, err := ioutil.ReadDir(dirPath)

	if err != nil {
		return
	}

	var orderList []entities.Order
	for _, file := range files {
		if strings.Index(file.Name(), ".xlsx") < 0 {
			continue
		}

		orders := utils.ExcelRead(dirPath + "/" + file.Name())

		if len(orders) > 0 {
			orderList = append(orderList, orders...)
		}
	}

	if len(orderList) > 0 {
		result := utils.ExcelCreate(dirPath, orderList)
		if result {
			fmt.Println("处理完成！")
		} else {
			fmt.Println("处理失败！")
		}
	} else {
		fmt.Println("excel中没有合适数据!")
	}
}
