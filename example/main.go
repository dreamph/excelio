package main

import (
	"fmt"
	"log"
	"time"

	"github.com/dreamph/excelio"
	"github.com/go-playground/validator/v10"
)

type Product struct {
	Code   string    `excel:"Code"   validate:"required"`
	Name   string    `col:"2"        validate:"required"`
	Price  float64   `excelcol:"C"   validate:"required,gt=0"`
	Active bool      `excel:"Active" validate:"required"`
	Since  time.Time `excel:"Since"  fmt:"2006-01-02"`
}

func main() {
	v := validator.New()

	products, rowErrs, err := excelio.ReadFile[Product](
		"products.xlsx",
		excelio.Sheet("Products"),
		excelio.Header(1),   // row 1 is header
		excelio.StartRow(2), // row 2 is first data row
		excelio.UseValidator(v),
	)
	if err != nil {
		log.Fatalf("read error: %v", err)
	}

	fmt.Println("== VALID ROWS ==")
	for i, p := range products {
		fmt.Printf("%d: %+v\n", i+1, p)
	}

	fmt.Println("== ROW ERRORS ==")
	for _, e := range rowErrs {
		fmt.Printf("row=%d col=%s field=%s msg=%v\n",
			e.ExcelRowIndex, e.ColLetter, e.Field, e.Err)
	}

	rowErrs, err = excelio.StreamFile[Product](
		"products.xlsx",
		excelio.Sheet("Products"),
		excelio.Header(1),
		excelio.StartRow(2),
		excelio.UseValidator(v),
		excelio.ErrCol(10), // error messages will be written into column J
		// per-row handler:
		excelio.OnStreamRow(func(rowIdx, logicalIdx int, p *Product, perRowErrs []excelio.RowError) error {
			if len(perRowErrs) > 0 {
				fmt.Printf("Row %d has %d error(s):\n", rowIdx, len(perRowErrs))
				for _, e := range perRowErrs {
					fmt.Printf("  col=%s field=%s msg=%v\n",
						e.ColLetter, e.Field, e.Err)
				}
				// continue processing other rows
				return nil
			}

			if p != nil {
				// This row is valid → you can insert into DB, etc.
				fmt.Printf("OK row=%d logical=%d product=%+v\n", rowIdx, logicalIdx, *p)
			}
			return nil
		}),
	)
	if err != nil {
		log.Fatalf("stream error: %v", err)
	}

	fmt.Println("== SUMMARY ERRORS ==")
	for _, e := range rowErrs {
		fmt.Printf("row=%d col=%s field=%s msg=%v\n",
			e.ExcelRowIndex, e.ColLetter, e.Field, e.Err)
	}
}
