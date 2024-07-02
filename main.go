package main

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

func main() {
	excel, err := excelize.OpenFile("repro.xlsx")
	if err != nil {
		panic(err)
	}

	cells := []string{
		"A1",
		"B1",
		"C1",
		"D1",
	}
	for _, cell := range cells {
		val, err := excel.CalcCellValue("Sheet1", cell)
		if err != nil {
			fmt.Printf("ERR %s: %s\n", cell, err.Error())
			continue
		}
		fmt.Printf("OK  %s: %s\n", cell, val)
	}
}
