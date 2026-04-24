package main

import (
	"fmt"
	"log"

	ods "github.com/mukbeast4/go-ods"
)

func main() {
	f := ods.NewFile()

	for i := 1; i <= 10; i++ {
		_ = f.SetCellFloat("Sheet1", fmt.Sprintf("A%d", i), float64(i))
	}

	f.SetCellFormula("Sheet1", "C1", "SUM([.A1:.A10])")
	f.SetCellFormula("Sheet1", "C2", "AVERAGE([.A1:.A10])")
	f.SetCellFormula("Sheet1", "C3", "MEDIAN([.A1:.A10])")
	f.SetCellFormula("Sheet1", "C4", "MAX([.A1:.A10])-MIN([.A1:.A10])")
	f.SetCellFormula("Sheet1", "C5", "IF([.C1]>50;\"big\";\"small\")")

	f.SetCellStr("Sheet1", "B1", "Sum")
	f.SetCellStr("Sheet1", "B2", "Average")
	f.SetCellStr("Sheet1", "B3", "Median")
	f.SetCellStr("Sheet1", "B4", "Range")
	f.SetCellStr("Sheet1", "B5", "Size")

	if err := f.RecalcAll(); err != nil {
		log.Fatalf("recalc: %v", err)
	}

	for _, ref := range []string{"C1", "C2", "C3", "C4"} {
		v, _ := f.GetCellFloat("Sheet1", ref)
		label, _ := f.GetCellValue("Sheet1", "B"+ref[1:])
		fmt.Printf("%-10s %v\n", label, v)
	}
	size, _ := f.GetCellValue("Sheet1", "C5")
	fmt.Printf("%-10s %s\n", "Size", size)

	if err := f.SaveAs("formulas.ods"); err != nil {
		log.Fatal(err)
	}
}
