package utils

import (
	"fmt"
	"math/rand"
	"strconv"
	"time"

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
		return false, "没有数据"
	}

	var file *xlsx.File
	var sheet *xlsx.Sheet
	var row *xlsx.Row
	var cell *xlsx.Cell
	var err error

	file = xlsx.NewFile()
	sheet, err = file.AddSheet("Sheet1")
	if err != nil {
		fmt.Printf("open failed:%s\n", err)
		return false
	}

	row = sheet.AddRow()
	row.SetHeightCM(1)
	cell = row.AddCell()
	cell.Value = "JOB NO"
	cell = row.AddCell()
	cell.Value = "QTY"
	cell = row.AddCell()
	cell.Value = "Item Code"
	cell = row.AddCell()
	cell.Value = "MM/YY"
	cell = row.AddCell()
	cell.Value = "Stock"
	cell = row.AddCell()
	cell.Value = "Type"
	cell = row.AddCell()
	cell.Value = "Sub"
	cell = row.AddCell()
	cell.Value = "Lot"
	cell = row.AddCell()
	cell.Value = "Line"
	cell = row.AddCell()
	cell.Value = "Size Code"
	cell = row.AddCell()
	cell.Value = "Description"
	cell = row.AddCell()
	cell.Value = "Brand Type"
	cell = row.AddCell()
	cell.Value = "Color"
	cell = row.AddCell()
	cell.Value = "Size"
	cell = row.AddCell()
	cell.Value = "Cat/Sku"
	cell = row.AddCell()
	cell.Value = "Product ID/Style# "
	cell = row.AddCell()
	cell.Value = "UPC"
	cell = row.AddCell()
	cell.Value = "128c"
	cell = row.AddCell()
	cell.Value = "Misc1"
	cell = row.AddCell()
	cell.Value = "Misc2"
	cell = row.AddCell()
	cell.Value = "2 or More"
	cell = row.AddCell()
	cell.Value = "Retail"
	cell = row.AddCell()
	cell.Value = "EPC Start"
	cell = row.AddCell()
	cell.Value = "EPC End"
	cell = row.AddCell()
	cell.Value = "Customer PO"
	cell = row.AddCell()
	cell.Value = "Country Of Origin"
	cell = row.AddCell()
	cell.Value = "Supplier #"
	cell = row.AddCell()
	cell.Value = "Location "
	cell = row.AddCell()
	cell.Value = "Location Code"
	cell = row.AddCell()
	cell.Value = "SpecialValue1"
	cell = row.AddCell()
	cell.Value = "SpecialValue2"
	cell = row.AddCell()
	cell.Value = "SpecialValue3"
	cell = row.AddCell()
	cell.Value = "SpecialValue4"
	cell = row.AddCell()
	cell.Value = "SpecialValue5"
	cell = row.AddCell()
	cell.Value = "SpecialValue6"
	cell = row.AddCell()
	cell.Value = "SpecialValue7"
	cell = row.AddCell()
	cell.Value = "SpecialValue8"
	cell = row.AddCell()
	cell.Value = "SpecialValue9"
	cell = row.AddCell()
	cell.Value = "SpecialValue10"

	for _, value := range orderList {

	}

	r := rand.New(rand.NewSource(time.Now().Unix()))
	xlsxName := "order" + (time.Now().Format("2006-01-02")) + strconv.Itoa(r.Intn(100)) + ".xlsx"

	err = file.Save(xlsxName)
	if err != nil {
		fmt.Printf("open failed:%s\n", err)
		return false
	}
	return true
}
