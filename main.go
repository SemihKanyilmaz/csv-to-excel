package main

import (
	"encoding/csv"
	"errors"
	"fmt"
	"os"
	"strconv"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
)

func main() {

	err := CsvToExcel("samplefile.csv")
	if err != nil {
		fmt.Println(err)
	}

}

func CsvToExcel(fileName string) error {

	fileSplits := strings.Split(fileName, ("."))

	if !strings.Contains(fileSplits[1], "csv") {
		return errors.New("The extension is not CSV")
	}

	const sheet string = "Sheet1"
	var letters = [...]string{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}

	src, err := os.Open(fileName)
	if err != nil {
		return err
	}
	defer src.Close()
	csvLines, err := csv.NewReader(src).ReadAll()
	if err != nil {
		return err
	}
	f := excelize.NewFile()
	ctr := 1

	for _, lines := range csvLines {
		for j, line := range lines {
			if len(line) > 0 && strings.TrimSpace(line) != "" {
				f.SetCellValue(sheet, letters[j]+strconv.Itoa(ctr), line)
			} else {
				f.SetCellValue(sheet, letters[j]+strconv.Itoa(ctr), "-")
			}
		}
		ctr++
	}
	newFileName := fileSplits[0] + ".xlsx"
	if err := f.SaveAs(newFileName); err != nil {
		return err
	}
	fmt.Printf("%s has been created successfully",newFileName)
	return nil
}
