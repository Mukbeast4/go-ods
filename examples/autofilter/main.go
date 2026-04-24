package main

import (
	"log"

	ods "github.com/mukbeast4/go-ods"
)

func main() {
	f := ods.NewFile()

	rows := [][]any{
		{"Region", "Product", "Units"},
		{"EU", "Widget", 120.0},
		{"EU", "Gadget", 60.0},
		{"US", "Widget", 80.0},
		{"US", "Gadget", 220.0},
		{"APAC", "Widget", 40.0},
	}
	if err := f.AppendRows("Sheet1", rows); err != nil {
		log.Fatal(err)
	}

	if err := f.SetAutoFilter("Sheet1", "A1", "C6"); err != nil {
		log.Fatal(err)
	}

	if err := f.SetFilterCriteria("Sheet1", []ods.FilterCriteria{
		{Column: 0, Values: []string{"EU", "US"}},
	}); err != nil {
		log.Fatal(err)
	}

	if err := f.SetSort("Sheet1", []ods.SortKey{
		{Column: 2, Descending: true},
	}); err != nil {
		log.Fatal(err)
	}

	if err := f.SaveAs("autofilter.ods"); err != nil {
		log.Fatal(err)
	}
}
