package utils

import (
	"fmt"
	"math/rand"
	"strconv"
	"time"

	"github.com/tealeg/xlsx"

	"Snai.Go.Excel/entities"
)

func ExcelRead(dirPath string) []entities.Order {
	var orderList []entities.Order
	fileName := dirPath
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

		orderQty := row.Cells[0].String()
		overageQty := row.Cells[1].String()
		totalQty := row.Cells[2].String()
		itemCode := row.Cells[3].String()
		mmyy := row.Cells[4].String()
		stock := row.Cells[5].String()
		typea := row.Cells[6].String()
		sub := row.Cells[7].String()
		lot := row.Cells[8].String()
		line := row.Cells[9].String()
		sizeCode := row.Cells[10].String()
		description := row.Cells[11].String()
		brandType := row.Cells[12].String()
		color := row.Cells[13].String()
		size := row.Cells[14].String()
		catSku := row.Cells[15].String()
		productIDStyle := row.Cells[16].String()
		uPC := row.Cells[17].String()
		c128 := row.Cells[18].String()
		misc1 := row.Cells[18].String()
		misc2 := row.Cells[20].String()
		moreOr2 := row.Cells[21].String()
		retail := row.Cells[22].String()
		ePCStart := row.Cells[23].String()
		ePCEnd := row.Cells[24].String()
		customerPO := row.Cells[25].String()
		countryOfOrigin := row.Cells[26].String()
		supplier := row.Cells[27].String()
		status := row.Cells[28].String()
		location := row.Cells[29].String()
		locationCode := row.Cells[30].String()
		specialValue1 := row.Cells[31].String()
		specialValue2 := row.Cells[32].String()
		specialValue3 := row.Cells[33].String()
		specialValue4 := row.Cells[34].String()
		specialValue5 := row.Cells[35].String()
		specialValue6 := row.Cells[36].String()
		specialValue7 := row.Cells[37].String()
		specialValue8 := row.Cells[38].String()
		specialValue9 := row.Cells[39].String()
		specialValue10 := row.Cells[40].String()

		o := entities.Order{JobNo: jobNo, OrderQty: orderQty, OverageQty: overageQty, TotalQty: totalQty, ItemCode: itemCode, MMYY: mmyy, Stock: stock, Type: typea, Sub: sub, Lot: lot, Line: line, SizeCode: sizeCode, Description: description, BrandType: brandType, Color: color, Size: size, CatSku: catSku, ProductIDStyle: productIDStyle, UPC: uPC, C128: c128, Misc1: misc1, Misc2: misc2, MoreOr2: moreOr2, Retail: retail, EPCStart: ePCStart, EPCEnd: ePCEnd, CustomerPO: customerPO, CountryOfOrigin: countryOfOrigin, Supplier: supplier, Status: status, Location: location, LocationCode: locationCode, SpecialValue1: specialValue1, SpecialValue2: specialValue2, SpecialValue3: specialValue3, SpecialValue4: specialValue4, SpecialValue5: specialValue5, SpecialValue6: specialValue6, SpecialValue7: specialValue7, SpecialValue8: specialValue8, SpecialValue9: specialValue9, SpecialValue10: specialValue10}

		orderList = append(orderList, o)
	}

	return orderList
}

func ExcelCreate(dirPath string, orders []entities.Order) bool {
	if len(orders) <= 0 {
		fmt.Printf("not data")
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
		fmt.Printf("addsheet failed:%s\n", err)
		return false
	}

	row = sheet.AddRow()
	row.SetHeightCM(1)
	cell = row.AddCell()
	cell.Value = "JOB NO"
	cell = row.AddCell()
	cell.Value = "Order Qty"
	cell = row.AddCell()
	cell.Value = "Overage Qty"
	cell = row.AddCell()
	cell.Value = "Total Qty"
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
	cell.Value = "status"
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

	for _, order := range orders {
		row = sheet.AddRow()
		row.SetHeightCM(1)
		cell = row.AddCell()
		cell.Value = order.JobNo
		cell = row.AddCell()
		cell.Value = order.OrderQty
		cell = row.AddCell()
		cell.Value = order.OverageQty
		cell = row.AddCell()
		cell.Value = order.TotalQty
		cell = row.AddCell()
		cell.Value = order.ItemCode
		cell = row.AddCell()
		cell.Value = order.MMYY
		cell = row.AddCell()
		cell.Value = order.Stock
		cell = row.AddCell()
		cell.Value = order.Type
		cell = row.AddCell()
		cell.Value = order.Sub
		cell = row.AddCell()
		cell.Value = order.Lot
		cell = row.AddCell()
		cell.Value = order.Line
		cell = row.AddCell()
		cell.Value = order.SizeCode
		cell = row.AddCell()
		cell.Value = order.Description
		cell = row.AddCell()
		cell.Value = order.BrandType
		cell = row.AddCell()
		cell.Value = order.Color
		cell = row.AddCell()
		cell.Value = order.Size
		cell = row.AddCell()
		cell.Value = order.CatSku
		cell = row.AddCell()
		cell.Value = order.ProductIDStyle
		cell = row.AddCell()
		cell.Value = order.UPC
		cell = row.AddCell()
		cell.Value = order.C128
		cell = row.AddCell()
		cell.Value = order.Misc1
		cell = row.AddCell()
		cell.Value = order.Misc2
		cell = row.AddCell()
		cell.Value = order.MoreOr2
		cell = row.AddCell()
		cell.Value = order.Retail
		cell = row.AddCell()
		cell.Value = order.EPCStart
		cell = row.AddCell()
		cell.Value = order.EPCEnd
		cell = row.AddCell()
		cell.Value = order.CustomerPO
		cell = row.AddCell()
		cell.Value = order.CountryOfOrigin
		cell = row.AddCell()
		cell.Value = order.Supplier
		cell = row.AddCell()
		cell.Value = order.Status
		cell = row.AddCell()
		cell.Value = order.Location
		cell = row.AddCell()
		cell.Value = order.LocationCode
		cell = row.AddCell()
		cell.Value = order.SpecialValue1
		cell = row.AddCell()
		cell.Value = order.SpecialValue2
		cell = row.AddCell()
		cell.Value = order.SpecialValue3
		cell = row.AddCell()
		cell.Value = order.SpecialValue4
		cell = row.AddCell()
		cell.Value = order.SpecialValue5
		cell = row.AddCell()
		cell.Value = order.SpecialValue6
		cell = row.AddCell()
		cell.Value = order.SpecialValue7
		cell = row.AddCell()
		cell.Value = order.SpecialValue8
		cell = row.AddCell()
		cell.Value = order.SpecialValue9
		cell = row.AddCell()
		cell.Value = order.SpecialValue10
	}

	r := rand.New(rand.NewSource(time.Now().Unix()))
	xlsxName := dirPath + "/" + "order" + (time.Now().Format("2006-01-02")) + "-" + strconv.Itoa(r.Intn(100)) + ".xlsx"

	err = file.Save(xlsxName)
	if err != nil {
		fmt.Printf("save failed:%s\n", err)
		return false
	}

	return true
}
