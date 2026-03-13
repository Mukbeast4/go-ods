//go:build ignore

package main

import (
	"log"

	ods "github.com/mukbeast4/go-ods"
)

func main() {
	f := ods.NewFile()

	f.SetCellStr("Sheet1", "A1", "Name")
	f.SetCellStr("Sheet1", "B1", "Score1")
	f.SetCellStr("Sheet1", "C1", "Score2")
	f.SetCellStr("Sheet1", "D1", "Score3")
	f.SetCellStr("Sheet1", "E1", "Total")
	f.SetCellStr("Sheet1", "F1", "Average")
	f.SetCellStr("Sheet1", "G1", "Status")

	names := []string{"Alice", "Bob", "Charlie", "Diana", "Eve"}
	scores := [][3]float64{
		{85, 92, 78},
		{70, 65, 80},
		{95, 88, 91},
		{60, 72, 68},
		{100, 95, 97},
	}

	for i, name := range names {
		row := i + 2
		rowStr := string(rune('0' + row))
		if row >= 10 {
			rowStr = "10"
		}
		_ = rowStr

		f.SetCellStr("Sheet1", cellRef("A", row), name)
		f.SetCellFloat("Sheet1", cellRef("B", row), scores[i][0])
		f.SetCellFloat("Sheet1", cellRef("C", row), scores[i][1])
		f.SetCellFloat("Sheet1", cellRef("D", row), scores[i][2])

		f.SetCellFormula("Sheet1", cellRef("E", row), "[.B"+itoa(row)+"]+[.C"+itoa(row)+"]+[.D"+itoa(row)+"]")
		f.SetCellFormula("Sheet1", cellRef("F", row), "([.B"+itoa(row)+"]+[.C"+itoa(row)+"]+[.D"+itoa(row)+"])/3")
		f.SetCellFormula("Sheet1", cellRef("G", row), "IF([.F"+itoa(row)+"]>=80;\"Pass\";\"Fail\")")
	}

	f.SetCellFormula("Sheet1", "B7", "SUM([.B2:.B6])")
	f.SetCellFormula("Sheet1", "C7", "SUM([.C2:.C6])")
	f.SetCellFormula("Sheet1", "D7", "SUM([.D2:.D6])")
	f.SetCellFormula("Sheet1", "E7", "SUM([.E2:.E6])")
	f.SetCellStr("Sheet1", "A7", "Total")

	f.SetCellStr("Sheet1", "A8", "Max")
	f.SetCellFormula("Sheet1", "E8", "MAX([.E2:.E6])")
	f.SetCellStr("Sheet1", "A9", "Min")
	f.SetCellFormula("Sheet1", "E9", "MIN([.E2:.E6])")

	f.RecalcAll()

	if err := f.SaveAs("testdata/sample.ods"); err != nil {
		log.Fatalf("SaveAs: %v", err)
	}
	log.Println("Generated testdata/sample.ods")
}

func cellRef(col string, row int) string {
	return col + itoa(row)
}

func itoa(i int) string {
	if i < 10 {
		return string(rune('0' + i))
	}
	return string(rune('0'+i/10)) + string(rune('0'+i%10))
}
