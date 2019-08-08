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

	fmt.Printf("len:%d", len(orderList))
	if len(orderList) > 0 {
		for _, value := range orderList {
			fmt.Printf("JobNo:%s,Qyt:%s", value.JobNo, value.Qyt)
		}
	}
}
