package entities

type order struct {
	JobNo           string
	Qyt             string
	ItemCode        string
	MMYY            string
	Stock           string
	Type            string
	Sub             string
	Lot             string
	Line            string
	SizeCode        string
	Description     string
	BrandType       string
	Color           string
	Size            string
	CatSku          string
	ProductIDStyle  string
	UPC             string
	C128            string
	Misc1           string
	Misc2           string
	MoreOr2         string
	Retail          string
	EPCStart        string
	EPCEnd          string
	CustomerPO      string
	CountryOfOrigin string
	Supplier        string
	Status          string
	Location        string
	LocationCode    string
	SpecialValue1   string
	SpecialValue2   string
	SpecialValue3   string
	SpecialValue4   string
	SpecialValue5   string
	SpecialValue6   string
	SpecialValue7   string
	SpecialValue8   string
	SpecialValue9   string
	SpecialValue10  string
}

func New(JobNo string,
	Qyt string,
	ItemCode string,
	MMYY string,
	Stock string,
	Type string,
	Sub string,
	Lot string,
	Line string,
	SizeCode string,
	Description string,
	BrandType string,
	Color string,
	Size string,
	CatSku string,
	ProductIDStyle string,
	UPC string,
	C128 string,
	Misc1 string,
	Misc2 string,
	MoreOr2 string,
	Retail string,
	EPCStart string,
	EPCEnd string,
	CustomerPO string,
	CountryOfOrigin string,
	Supplier string,
	Status string,
	Location string,
	LocationCode string,
	SpecialValue1 string,
	SpecialValue2 string,
	SpecialValue3 string,
	SpecialValue4 string,
	SpecialValue5 string,
	SpecialValue6 string,
	SpecialValue7 string,
	SpecialValue8 string,
	SpecialValue9 string,
	SpecialValue10 string) order {
	o := order{JobNo,
		Qyt,
		ItemCode,
		MMYY,
		Stock,
		Type,
		Sub,
		Lot,
		Line,
		SizeCode,
		Description,
		BrandType,
		Color,
		Size,
		CatSku,
		ProductIDStyle,
		UPC,
		C128,
		Misc1,
		Misc2,
		MoreOr2,
		Retail,
		EPCStart,
		EPCEnd,
		CustomerPO,
		CountryOfOrigin,
		Supplier,
		Status,
		Location,
		LocationCode,
		SpecialValue1,
		SpecialValue2,
		SpecialValue3,
		SpecialValue4,
		SpecialValue5,
		SpecialValue6,
		SpecialValue7,
		SpecialValue8,
		SpecialValue9,
		SpecialValue10}
	return o
}
