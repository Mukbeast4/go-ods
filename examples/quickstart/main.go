package main

import (
	"fmt"
	"log"

	ods "github.com/mukbeast4/go-ods"
)

func main() {
	f := ods.NewFile()

	f.SetCellStr("Sheet1", "A1", "Product")
	f.SetCellStr("Sheet1", "B1", "Price")
	f.SetCellStr("Sheet1", "A2", "Widget")
	f.SetCellFloat("Sheet1", "B2", 19.99)
	f.SetCellStr("Sheet1", "A3", "Gadget")
	f.SetCellFloat("Sheet1", "B3", 42.50)
	f.SetCellStr("Sheet1", "A4", "Total")
	f.SetCellFormula("Sheet1", "B4", "SUM([.B2:.B3])")

	if err := f.RecalcSheet("Sheet1"); err != nil {
		log.Fatalf("recalc: %v", err)
	}

	total, _ := f.GetCellFloat("Sheet1", "B4")
	fmt.Printf("total = %.2f\n", total)

	out := "quickstart.ods"
	if err := f.SaveAs(out); err != nil {
		log.Fatalf("save: %v", err)
	}
	fmt.Println("wrote", out)
}
