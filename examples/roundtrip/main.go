package main

import (
	"fmt"
	"log"
	"os"

	ods "github.com/mukbeast4/go-ods"
)

func main() {
	path := "roundtrip.ods"

	f := ods.NewFile()
	f.SetCellStr("Sheet1", "A1", "counter")
	f.SetCellInt("Sheet1", "B1", 0)
	if err := f.SaveAs(path); err != nil {
		log.Fatal(err)
	}

	reopened, err := ods.OpenFile(path)
	if err != nil {
		log.Fatal(err)
	}

	current, _ := reopened.GetCellInt("Sheet1", "B1")
	reopened.SetCellInt("Sheet1", "B1", current+1)

	if err := reopened.SaveAs(path); err != nil {
		log.Fatal(err)
	}

	final, err := ods.OpenFile(path)
	if err != nil {
		log.Fatal(err)
	}
	got, _ := final.GetCellInt("Sheet1", "B1")
	fmt.Printf("counter after round-trip = %d\n", got)

	if err := os.Remove(path); err != nil {
		log.Printf("cleanup: %v", err)
	}
}
