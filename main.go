package main

import (
	"fmt"

	"Snai.Go.Excel/entities"
	"Snai.Go.Excel/utils"
)

func main(){
	files, err := ioutil.ReadDir("excel")

	if err != nil{
        return
	}
	
	for _, file := range files{
        if !strings.Contains(file.Name, ".xls") && !strings.Contains(file.Name, ".xlsx"){
            continue
		}
		
    }
}