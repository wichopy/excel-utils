package main

import (
	"fmt"
	"log"
	"os"

	"github.com/360EntSecGroup-Skylar/excelize"
)

func main() {
	output := excelize.NewFile()
	ignoreTopRow := true
	targetSheet := "Some-data"
	outputRow := 0
	// Open all files in a directory
	// X Grab alldata from a sheet.
	// X Combine sheet to make a single dataset

	fmt.Println("Opening fileBook1.xlsx")
	f, err := excelize.OpenFile("Book1.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}

	// Get all the rows in the Sheet1.
	fmt.Println("Copying data from sheet" + targetSheet)
	rows, err := f.GetRows(targetSheet)
	for i, row := range rows {
		if ignoreTopRow && i == 0 {
			continue
		}
		for j, colCell := range row {
			cellCoords, err := excelize.CoordinatesToCellName(j+1, outputRow+1)
			if err != nil {
				fmt.Println(err)
				return
			}
			output.SetCellValue("Sheet1", cellCoords, colCell)
			// fmt.Print(colCell, "\t")
		}
		// fmt.Println()
		outputRow += 1
		// if i > 20 {
		// 	break
		// }
	}

	fmt.Printf("%v%s\n", outputRow, "rows copied")

	dir, err := os.Getwd()
	if err != nil {
		log.Fatal(err)
	}
	fmt.Println("Saving to " + dir + "/output.xlsx")
	output.SaveAs(dir + "/output.xlsx")

}
