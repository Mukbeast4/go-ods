package main

import (
	"fmt"
	"log"

	ods "github.com/mukbeast4/go-ods"
)

func main() {
	f := ods.NewFile()

	const rows = 10000
	batch := make([][]any, 0, 1000)
	for i := 1; i <= rows; i++ {
		batch = append(batch, []any{i, float64(i) * 1.5, fmt.Sprintf("row-%d", i)})
		if len(batch) == 1000 {
			if err := f.AppendRows("Sheet1", batch); err != nil {
				log.Fatal(err)
			}
			batch = batch[:0]
		}
	}
	if len(batch) > 0 {
		if err := f.AppendRows("Sheet1", batch); err != nil {
			log.Fatal(err)
		}
	}

	it, err := f.NewRowIterator("Sheet1")
	if err != nil {
		log.Fatal(err)
	}

	var seen, last int
	for it.Next() {
		seen++
		last = it.RowIndex()
	}
	fmt.Printf("iterated %d rows, last index = %d\n", seen, last)

	if err := f.SaveAs("streaming.ods"); err != nil {
		log.Fatal(err)
	}
}
