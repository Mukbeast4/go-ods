package main

import (
	"log"

	ods "github.com/mukbeast4/go-ods"
)

func main() {
	f := ods.NewFile()

	header, _ := f.NewStyle(&ods.Style{
		Font: &ods.Font{
			Family: "Arial",
			Size:   "12pt",
			Bold:   "bold",
			Color:  "#FFFFFF",
		},
		Fill: &ods.Fill{Color: "#336699"},
		Alignment: &ods.Alignment{
			Horizontal: "center",
			Vertical:   "middle",
		},
		Border: &ods.Border{Style: "solid", Width: "0.5pt", Color: "#000000"},
	})

	odd, _ := f.NewStyle(&ods.Style{Fill: &ods.Fill{Color: "#F5F5F5"}})

	rotated, _ := f.NewStyle(&ods.Style{
		Alignment: &ods.Alignment{Rotation: 45},
	})

	f.SetCellStr("Sheet1", "A1", "Name")
	f.SetCellStr("Sheet1", "B1", "Score")
	f.SetCellStr("Sheet1", "C1", "Grade")
	f.SetCellStyle("Sheet1", "A1", "C1", header)

	f.Range("Sheet1", "A2:C2").SetValue("Alice").SetStyle(odd)
	f.SetCellStr("Sheet1", "A3", "Bob")
	f.SetCellFloat("Sheet1", "B3", 88)
	f.SetCellStr("Sheet1", "C3", "B+")

	f.SetCellStr("Sheet1", "E1", "Rotated header")
	f.SetCellStyle("Sheet1", "E1", "E1", rotated)

	f.SetRowHeight("Sheet1", 1, 1.0)

	if err := f.SaveAs("styling.ods"); err != nil {
		log.Fatal(err)
	}
}
