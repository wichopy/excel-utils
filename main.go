package main

import (
	"flag"
	"fmt"
	"io/ioutil"
	"log"
	"os"
	"strings"
	"path"

	"github.com/360EntSecGroup-Skylar/excelize"
)

func main() {
	output := excelize.NewFile()
	var ignoreTopRow = flag.Bool("ignoreTopRow", true, "Ignore top row in all sheets")
	var targetSheet = flag.String("targetSheet", "Sheet1", "Sheet to target in each workbook")
	var targetDir = flag.String("targetDirectory", ".", "Path to directory for scanning.")

	flag.Parse()
	outputRow := 0
	fileExtFilter := "xlsx"

	fmt.Printf("Starting excel merger with flags: %v: %v, %v: %v, %v:%v \n", "ignoreTopRow", *ignoreTopRow, "targetSheet", *targetSheet, "targetDirectory", *targetDir)
	fmt.Println()
	pwd, err := os.Getwd()
	if err != nil {
		log.Fatal(err)
		return
	}

	fmt.Println("Scanning files in target directory: '" + *targetDir + "'")
	fmt.Println()
	files, err := ioutil.ReadDir(*targetDir)
	if err != nil {
		log.Fatal(err)
	}

	for _, file := range files {
		if !strings.Contains(file.Name(), fileExtFilter) {
			continue
		}

		fmt.Println("Opening " + file.Name())
		f, err := excelize.OpenFile(path.Join(*targetDir, file.Name()))
		if err != nil {
			fmt.Println(err)
			return
		}

		// Get all the rows in targetSheet.
		rows, err := f.GetRows(*targetSheet)
		if err != nil {
			fmt.Println("Cannot find target sheet '" + *targetSheet + "' in file '" + file.Name() + "'")
			continue
		}

		fmt.Println("Copying data from sheet " + *targetSheet + " to output.")

		for i, row := range rows {
			if *ignoreTopRow && i == 0 {
				continue
			}

			for j, colCell := range row {
				cellCoords, err := excelize.CoordinatesToCellName(j+1, outputRow+1)
				if err != nil {
					fmt.Println(err)
					return
				}
				output.SetCellValue("Sheet1", cellCoords, colCell)
			}

			outputRow += 1
		}
	}

	fmt.Println()
	fmt.Printf("%v%s\n", outputRow, " rows copied\n")

	outputLocation := path.Join(pwd, "output.xlsx")
	fmt.Println("Saving to " + outputLocation)
	output.SaveAs(outputLocation)
}
