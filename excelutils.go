package utils

import (
	"fmt"

	"github.com/tealeg/xlsx"

	"Snai.Go.Excel/entities"
)

func excelread(string path) []order {
	var orderList []order
	fileName := path
	xlFile, err := xlsx.OpenFile(fileName)
	if err != nil {
		fmt.Printf("open failed:%s\n", err)
		return orderList
	}

	sheet := xlFile.Sheets[0]

	jobNo := sheet.Rows[0].Cells[1]
	
	for index, value := range sheet.Rows {
		if index < 14 {
			continue 
		}

		qyt := row.Cells[0]
		itemCode := row.Cells[1]
		mmyy := row.Cells[2]
		stock := row.Cells[3]
		typea := row.Cells[4]
		sub := row.Cells[5]
		lot := row.Cells[6]
		line := row.Cells[7]
		sizeCode := row.Cells[8]
		description := row.Cells[9]
		brandType := row.Cells[10]
		color := row.Cells[11]
		size := row.Cells[12]
		catSku := row.Cells[13]
		productIDStyle := row.Cells[14]
		uPC := row.Cells[15]
		c128 := row.Cells[16]
		misc1 := row.Cells[17]
		misc2 := row.Cells[18]
		moreOr2 := row.Cells[19]
		retail := row.Cells[20]
		ePCStart := row.Cells[21]
		ePCEnd := row.Cells[22]
		customerPO := row.Cells[23]
		countryOfOrigin := row.Cells[24]
		supplier := row.Cells[25]
		status := row.Cells[26]
		location := row.Cells[27]
		locationCode := row.Cells[28]
		specialValue1 := row.Cells[29]
		specialValue2 := row.Cells[30]
		specialValue3 := row.Cells[31]
		specialValue4 := row.Cells[32]
		specialValue5 := row.Cells[33]
		specialValue6 := row.Cells[34]
		specialValue7 := row.Cells[35]
		specialValue8 := row.Cells[36]
		specialValue9 := row.Cells[37]
		specialValue10 := row.Cells[38]

		o := order.New(jobNo,qyt,itemCode,mmyy,stock,typea,sub,lot,line,sizeCode,description,brandType,color,size,catSku,productIDStyle,uPC,c128,misc1,misc2,moreOr2,retail,ePCStart,ePCEnd,customerPO,countryOfOrigin,supplier,status,location,locationCode,specialValue1,specialValue2,specialValue3,specialValue4,specialValue5,specialValue6,specialValue7,specialValue8,specialValue9,specialValue10)

		orderList = append(orderList, o)
	}

	return orderList
}
