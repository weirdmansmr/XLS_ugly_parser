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

func processRow(rowHtml *goquery.Selection, startRow int, sheetName string, f *excelize.File, filePath string, startRowMap map[string]int) map[string]int {
	var sums [4]float64
	var values [4]float64
	coefficients := [4]float64{5, 4, 4, 5}
	weightCoefficients := [4]float64{5, 4, 3, 2}

	rowHtml.Find("td").Each(func(colIndex int, colHtml *goquery.Selection) {
		if colIndex >= 1 && colIndex <= 5 {
			cellValue := colHtml.Text()
			if num, err := strconv.ParseFloat(cellValue, 64); err == nil {
				if colIndex >= 2 && colIndex <= 5 {
					sums[colIndex-2] += num * coefficients[colIndex-2]
					values[colIndex-2] = num
				}
			}
			cell := fmt.Sprintf("%s%d", string('A'+colIndex), startRow)
			f.SetCellValue(sheetName, cell, cellValue)
			
			summaryCell := fmt.Sprintf("%s%d", string('A'+colIndex), startRowMap["Сводка"])
			f.SetCellValue("Сводка", summaryCell, cellValue)
		}
	})

	for i, sum := range sums {
		cell := fmt.Sprintf("%s%d", string('H'+i), startRow)
		f.SetCellValue(sheetName, cell, sum)
		
		summaryCell := fmt.Sprintf("%s%d", string('H'+i), startRowMap["Сводка"])
		f.SetCellValue("Сводка", summaryCell, sum)
	}

	weightedSum := values[0]*weightCoefficients[0] + values[1]*weightCoefficients[1] + values[2]*weightCoefficients[2] + values[3]*weightCoefficients[3]
	totalValues := values[0] + values[1] + values[2] + values[3]
	
	var avgGrade float64
	if totalValues > 0 {
		avgGrade = weightedSum / totalValues
	}
	
	f.SetCellValue(sheetName, fmt.Sprintf("M%d", startRow), avgGrade)
	f.SetCellValue("Сводка", fmt.Sprintf("M%d", startRowMap["Сводка"]), avgGrade)


	totalSum := sums[0] + sums[1] - sums[2] - sums[3]
	f.SetCellValue(sheetName, fmt.Sprintf("L%d", startRow), totalSum)
	f.SetCellValue("Сводка", fmt.Sprintf("L%d", startRowMap["Сводка"]), totalSum)

	var finalGrade float64
	
	onlyDFilled := values[1] > 0 && values[0] == 0 && values[2] == 0 && values[3] == 0
	
	if onlyDFilled {
		finalGrade = 0
	} else if avgGrade > 0 {
		finalGrade = totalSum / avgGrade
	}
	
	f.SetCellValue(sheetName, fmt.Sprintf("O%d", startRow), finalGrade)
	f.SetCellValue("Сводка", fmt.Sprintf("O%d", startRowMap["Сводка"]), finalGrade)

	f.SetCellValue(sheetName, fmt.Sprintf("G%d", startRow), strings.TrimSuffix(filePath, ".xls"))
	f.SetCellValue("Сводка", fmt.Sprintf("G%d", startRowMap["Сводка"]), strings.TrimSuffix(filePath, ".xls"))

	startRowMap["Сводка"]++
	
	return startRowMap
}

func processFile(filePath string, f *excelize.File, startRow int, sheetName string, startRowMap map[string]int) int {
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
				startRowMap = processRow(rowHtml, startRow, sheetName, f, filePath, startRowMap)
				startRow++
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

func sortSheetByColumn(f *excelize.File, sheetName string, lastRow int) {
	if lastRow <= 2 {
		return
	}

	dataRange := fmt.Sprintf("A2:O%d", lastRow-1)
	
	err := f.AutoFilter(sheetName, dataRange, []excelize.AutoFilterOptions{
		{
			Column:     "O",
			Expression: "",
		},
	})
	if err != nil {
		log.Printf("Ошибка создания автофильтра для %s: %v", sheetName, err)
		return
	}

	rows, err := f.GetRows(sheetName)
	if err != nil {
		log.Printf("Ошибка получения строк для %s: %v", sheetName, err)
		return
	}

	if len(rows) <= 2 {
		return
	}

	dataRows := rows[1:]

	for i := 0; i < len(dataRows)-1; i++ {
		for j := 0; j < len(dataRows)-i-1; j++ {
			var val1, val2 float64
			
			if len(dataRows[j]) > 14 {
				if v, err := strconv.ParseFloat(dataRows[j][14], 64); err == nil {
					val1 = v
				}
			}
			if len(dataRows[j+1]) > 14 {
				if v, err := strconv.ParseFloat(dataRows[j+1][14], 64); err == nil {
					val2 = v
				}
			}
			
			if val1 < val2 {
				dataRows[j], dataRows[j+1] = dataRows[j+1], dataRows[j]
			}
		}
	}

	for rowIdx, row := range dataRows {
		for colIdx, cellValue := range row {
			cell := fmt.Sprintf("%s%d", string('A'+colIdx), rowIdx+2)
			f.SetCellValue(sheetName, cell, cellValue)
		}
	}
}

func main() {
	f := excelize.NewFile()

	f.NewSheet("Сводка")
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
		"Б1.23-31.xls",
		"БАС1.24-21.xls",
		"БАС1.25-11.xls",
		"И1.22-41.xls",
		"И1.22-42.xls",
		"И1.22-43.xls",
		"И1.22-44.xls",
		"И1.23-31.xls",
		"И1.23-32.xls",
		"И1.23-33.xls",
		"И1.24-21.xls",
		"И1.24-22.xls",
		"И1.24-23.xls",
		"И1.25-11.xls",
		"И1.25-12.xls",
		"И1.25-13.xls",
		"И1.25-14.xls",
		"И2.22-41.xls",
		"И2.23-31.xls",
		"И2.24-21.xls",
		"И2.25-11.xls",
		"И2.25-12.xls",
		"ИБ1.24-21.xls",
		"ИБ1.25-11.xls",
		"О1.24-21.xls",
		"О1.24-22.xls",
		"О1.24-23.xls",
		"О1.25-11.xls",
		"О1.25-12.xls",
		"О1.25-13.xls",
		"О2.24-21.xls",
		"О2.25-11.xls",
		"О2.25-12.xls",
		"О5.25-11.xls",
		"О6.25-11.xls",
		"О6.25-12.xls",
		"С1.22-41.xls",
		"С1.23-31.xls",
		"С1.24-21.xls",
		"С1.25-11.xls",
		"СР1.23-31.xls",
		"СР1.23-32.xls",
		"СР1.24-21.xls",
		"СР1.24-22.xls",
		"СР1.25-11.xls",
		"Э1.23-31.xls",
		"Э1.23-32.xls",
		"Э1.24-21.xls",
		"Э1.24-22.xls",
		"Э1.25-11.xls",
		"Э1.25-12.xls",
		"Э3.23-31.xls",
	}

	startRowMap := map[string]int{
		"Сводка":   2,
		"4 курс":   2,
		"3 курс":   2,
		"2 курс":   2,
		"1 курс":   2,
	}

	for _, filePath := range files {
		groupNumber, err := getGroupNumber(filePath)
		if err != nil {
			log.Printf("Ошибка при обработке имени файла: %v", err)
			continue
		}

		var sheetName string
		switch groupNumber {
		case "22":
			sheetName = "4 курс"
		case "23":
			sheetName = "3 курс"
		case "24":
			sheetName = "2 курс"
		case "25":
			sheetName = "1 курс"
		default:
			log.Printf("Пропускаем файл с номером группы %s: %s", groupNumber, filePath)
			continue
		}

		startRowMap[sheetName] = processFile(filePath, f, startRowMap[sheetName], sheetName, startRowMap)
	}

	sheets := []string{"Сводка", "4 курс", "3 курс", "2 курс", "1 курс"}
	for _, sheet := range sheets {
		sortSheetByColumn(f, sheet, startRowMap[sheet])
	}
}
