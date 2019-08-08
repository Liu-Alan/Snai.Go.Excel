package utils

import (
	"fmt"

	"github.com/tealeg/xlsx"

	"Snai.Go.Excel/entities"
)

func ExcelRead(path string) []entities.Order {
	var orderList []entities.Order
	fileName := path
	xlFile, err := xlsx.OpenFile(fileName)
	if err != nil {
		fmt.Printf("open failed:%s\n", err)
		return orderList
	}

	sheet := xlFile.Sheets[0]

	jobNo := sheet.Rows[0].Cells[1].String()

	for index, row := range sheet.Rows {
		if index < 14 {
			continue
		}

		qyt := row.Cells[0].String()
		itemCode := row.Cells[1].String()
		mmyy := row.Cells[2].String()
		stock := row.Cells[3].String()
		typea := row.Cells[4].String()
		sub := row.Cells[5].String()
		lot := row.Cells[6].String()
		line := row.Cells[7].String()
		sizeCode := row.Cells[8].String()
		description := row.Cells[9].String()
		brandType := row.Cells[10].String()
		color := row.Cells[11].String()
		size := row.Cells[12].String()
		catSku := row.Cells[13].String()
		productIDStyle := row.Cells[14].String()
		uPC := row.Cells[15].String()
		c128 := row.Cells[16].String()
		misc1 := row.Cells[17].String()
		misc2 := row.Cells[18].String()
		moreOr2 := row.Cells[19].String()
		retail := row.Cells[20].String()
		ePCStart := row.Cells[21].String()
		ePCEnd := row.Cells[22].String()
		customerPO := row.Cells[23].String()
		countryOfOrigin := row.Cells[24].String()
		supplier := row.Cells[25].String()
		status := row.Cells[26].String()
		location := row.Cells[27].String()
		locationCode := row.Cells[28].String()
		specialValue1 := row.Cells[29].String()
		specialValue2 := row.Cells[30].String()
		specialValue3 := row.Cells[31].String()
		specialValue4 := row.Cells[32].String()
		specialValue5 := row.Cells[33].String()
		specialValue6 := row.Cells[34].String()
		specialValue7 := row.Cells[35].String()
		specialValue8 := row.Cells[36].String()
		specialValue9 := row.Cells[37].String()
		specialValue10 := row.Cells[38].String()

		o := entities.Order{JobNo: jobNo, Qyt: qyt, ItemCode: itemCode, MMYY: mmyy, Stock: stock, Type: typea, Sub: sub, Lot: lot, Line: line, SizeCode: sizeCode, Description: description, BrandType: brandType, Color: color, Size: size, CatSku: catSku, ProductIDStyle: productIDStyle, UPC: uPC, C128: c128, Misc1: misc1, Misc2: misc2, MoreOr2: moreOr2, Retail: retail, EPCStart: ePCStart, EPCEnd: ePCEnd, CustomerPO: customerPO, CountryOfOrigin: countryOfOrigin, Supplier: supplier, Status: status, Location: location, LocationCode: locationCode, SpecialValue1: specialValue1, SpecialValue2: specialValue2, SpecialValue3: specialValue3, SpecialValue4: specialValue4, SpecialValue5: specialValue5, SpecialValue6: specialValue6, SpecialValue7: specialValue7, SpecialValue8: specialValue8, SpecialValue9: specialValue9, SpecialValue10: specialValue10}

		orderList = append(orderList, o)
	}

	return orderList
}

func ExcelCreate(orders []entities.Order) bool {
	if len(orders) <= 0 {
		return false
	}

	var file *xlsx.File
	var sheet *xlsx.Sheet
	var row *xlsx.Row
	var cell *xlsx.Cell
	var err error

	file = xlsx.NewFile()
	sheet, err = file.AddSheet("Sheet1")
	if err != nil {
		fmt.Printf("create failed:%s\n", err)
		return false
	}

	row = sheet.AddRow()
	row.SetHeightCM(1)
	cell = row.AddCell()
	cell.Value = "JOB NO"

	return true
}
