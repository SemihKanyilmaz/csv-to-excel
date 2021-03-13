package main

import (
	"encoding/csv"
	"errors"
	"log"
	"mime/multipart"
	"net/http"
	"strconv"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/labstack/echo"
)

func main() {
	e := echo.New()
	e.HideBanner = true

	e.POST("/", Upload)

	log.Fatal(e.Start(":1923"))
}

func CsvToExcel(file *multipart.FileHeader) (string, error) {

	fileSplits := strings.Split(file.Filename, ("."))

	if !strings.Contains(fileSplits[1], "csv") {
		return "", errors.New("The extension is not CSV")
	}

	const sheet string = "Sheet1"
	var letters = [...]string{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}
	src, err := file.Open()
	if err != nil {
		return "", errors.New("An error occurred while opening the file")
	}
	defer src.Close()
	csvLines, err := csv.NewReader(src).ReadAll()
	if err != nil {
		return "", errors.New("An error occurred while reading the file")
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
		return "", errors.New("An error occurred while creating the file")
	}
	return newFileName, nil
}

func Upload(c echo.Context) error {

	f, err := c.FormFile("file")

	if err != nil {
		return c.NoContent(http.StatusBadRequest)
	}

	file, err := CsvToExcel(f)

	if err != nil {
		return c.JSON(500, err.Error())
	}

	return c.Attachment(file, file)
}
