package main

import (
	"fmt"
	"log"
	"os"
	"strconv"
	"strings"

	"github.com/PuerkitoBio/goquery"
	"github.com/xuri/excelize/v2"
)

func processRow(rowHtml *goquery.Selection, startRow int, sheetName string, f *excelize.File, filePath string) int {
	var sums [4]float64
	coefficients := [4]float64{5, 4, 4, 5}

	rowHtml.Find("td").Each(func(colIndex int, colHtml *goquery.Selection) {
		if colIndex >= 1 && colIndex <= 5 {
			cellValue := colHtml.Text()
			if num, err := strconv.ParseFloat(cellValue, 64); err == nil {
				if colIndex >= 2 && colIndex <= 5 {
					sums[colIndex-2] += num * coefficients[colIndex-2]
				}
			}
			cell := fmt.Sprintf("%s%d", string('A'+colIndex), startRow)
			f.SetCellValue(sheetName, cell, cellValue)
		}
	})

	for i, sum := range sums {
		cell := fmt.Sprintf("%s%d", string('H'+i), startRow)
		f.SetCellValue(sheetName, cell, sum)
	}

	totalSum := sums[0] + sums[1] - sums[2] - sums[3]
	f.SetCellValue(sheetName, fmt.Sprintf("L%d", startRow), totalSum)

	f.SetCellValue(sheetName, fmt.Sprintf("G%d", startRow), strings.TrimSuffix(filePath, ".xls"))

	return startRow + 1
}

func processFile(filePath string, f *excelize.File, startRow int, sheetName string) int {
	file, err := os.Open(filePath)
	if err != nil {
		log.Printf("Ошибка открытия файла %s: %v", filePath, err)
		return startRow
	}
	defer file.Close()

	doc, err := goquery.NewDocumentFromReader(file)
	if err != nil {
		log.Printf("Ошибка чтения файла %s как HTML: %v", filePath, err)
		return startRow
	}

	doc.Find("table").Each(func(tableIndex int, tableHtml *goquery.Selection) {
		tableHtml.Find("tr").Each(func(rowIndex int, rowHtml *goquery.Selection) {
			if rowIndex >= 1 {
				startRow = processRow(rowHtml, startRow, sheetName, f, filePath)
			}
		})
	})

	return startRow
}

func getGroupNumber(filePath string) (string, error) {
	index := strings.Index(filePath, ".")
	if index == -1 || len(filePath) <= index+2 {
		return "", fmt.Errorf("некорректный формат имени файла: %s", filePath)
	}

	return filePath[index+1 : index+3], nil
}

func main() {
	f := excelize.NewFile()

	f.NewSheet("4 курс")
	f.NewSheet("3 курс")
	f.NewSheet("2 курс")
	f.NewSheet("1 курс")

	defer func() {
		if err := f.SaveAs("result.xlsx"); err != nil {
			log.Fatalf("Ошибка сохранения файла: %v", err)
		}
	}()

	files := []string{
		"Б1.22-31.xls",
		"Б1.22-32.xls",
		"Б1.23-21.xls",
		"Б1.23-22.xls",
		"БАС1.24-11.xls",
		"И1.22-31.xls",
		"И1.22-32.xls",
		"И1.22-33.xls",
		"И1.22-34.xls",
		"И1.23-21.xls",
		"И1.23-22.xls",
		"И1.23-23.xls",
		"И1.24-11.xls",
		"И1.24-12.xls",
		"И1.24-13.xls",
		"И2.22-31.xls",
		"И2.22-32.xls",
		"И2.23-21.xls",
		"И2.24-11.xls",
		"ИБ1.24-11.xls",
		"О1.24-11.xls",
		"О1.24-12.xls",
		"О1.24-13.xls",
		"О2.24-11.xls",
		"О5.24-11.xls",
		"О5.24-12.xls",
		"О6.24-11.xls",
		"С1.21-41.xls",
		"С1.22-31.xls",
		"С1.23-21.xls",
		"С1.24-11.xls",
		"СР1.22-31.xls",
		"СР1.22-32.xls",
		"СР1.22-33.xls",
		"СР1.23-21.xls",
		"СР1.23-22.xls",
		"СР1.24-11.xls",
		"СР1.24-12.xls",
		"Э1.22-31.xls",
		"Э1.22-32.xls",
		"Э1.22-33.xls",
		"Э1.22-34.xls",
		"Э1.23-21.xls",
		"Э1.23-22.xls",
		"Э1.24-11.xls",
		"Э1.24-12.xls",
		"Э3.23-21.xls",
	}

	startRowMap := map[string]int{
		"Sheet_21": 2,
		"Sheet_22": 2,
		"Sheet_23": 2,
		"Sheet_24": 2,
	}

	for _, filePath := range files {
		groupNumber, err := getGroupNumber(filePath)
		if err != nil {
			log.Printf("Ошибка при обработке имени файла: %v", err)
			continue
		}

		var sheetName string
		switch groupNumber {
		case "21":
			sheetName = "4 курс"
		case "22":
			sheetName = "3 курс"
		case "23":
			sheetName = "2 курс"
		case "24":
			sheetName = "1 курс"
		default:
			log.Printf("Пропускаем файл с номером группы %s: %s", groupNumber, filePath)
			continue
		}

		startRowMap[sheetName] = processFile(filePath, f, startRowMap[sheetName], sheetName)
	}
}
