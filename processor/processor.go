package processor

import (
	"XLS_ugly_parser/calculator"
	"fmt"
	"log"
	"os"
	"strconv"
	"strings"

	"github.com/PuerkitoBio/goquery"
	"github.com/xuri/excelize/v2"
)

// Processor отвечает за обработку файлов и строк
type Processor struct {
	calc *calculator.Calculator
}

// NewProcessor создает новый экземпляр процессора
func NewProcessor() *Processor {
	return &Processor{
		calc: calculator.NewCalculator(),
	}
}

// ProcessFile обрабатывает один XLS файл
func (p *Processor) ProcessFile(filePath string, f *excelize.File, startRow int, sheetName string, startRowMap map[string]int) int {
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
				startRowMap = p.processRow(rowHtml, startRow, sheetName, f, filePath, startRowMap)
				startRow++
			}
		})
	})

	return startRow
}

// processRow обрабатывает одну строку таблицы
func (p *Processor) processRow(rowHtml *goquery.Selection, startRow int, sheetName string, f *excelize.File, filePath string, startRowMap map[string]int) map[string]int {
	var values [4]float64

	// Читаем данные из ячеек
	rowHtml.Find("td").Each(func(colIndex int, colHtml *goquery.Selection) {
		if colIndex >= 1 && colIndex <= 5 {
			cellValue := colHtml.Text()
			
			// Парсим числовые значения
			if num, err := strconv.ParseFloat(cellValue, 64); err == nil {
				if colIndex >= 2 && colIndex <= 5 {
					values[colIndex-2] = num
				}
			}
			
			// Записываем исходное значение в оба листа
			p.setCellValue(f, sheetName, startRow, colIndex, cellValue)
			p.setCellValue(f, "Сводка", startRowMap["Сводка"], colIndex, cellValue)
		}
	})

	// Вычисляем все метрики
	data := p.calc.CalculateRowData(values)

	// Записываем результаты вычислений
	p.writeResults(f, sheetName, startRow, startRowMap["Сводка"], data, filePath)

	// Увеличиваем счетчик для сводной страницы
	startRowMap["Сводка"]++

	return startRowMap
}

// setCellValue устанавливает значение ячейки
func (p *Processor) setCellValue(f *excelize.File, sheetName string, row, col int, value interface{}) {
	cell := fmt.Sprintf("%s%d", string('A'+col), row)
	f.SetCellValue(sheetName, cell, value)
}

// writeResults записывает все вычисленные результаты в Excel
func (p *Processor) writeResults(f *excelize.File, sheetName string, courseRow, summaryRow int, data *calculator.RowData, filePath string) {
	// Записываем суммы (столбцы H, I, J, K)
	for i, sum := range data.Sums {
		p.setCellValue(f, sheetName, courseRow, 'H'-'A'+i, sum)
		p.setCellValue(f, "Сводка", summaryRow, 'H'-'A'+i, sum)
	}

	// Записываем средневзвешенную оценку (столбец M)
	p.setCellValue(f, sheetName, courseRow, 'M'-'A', data.AvgGrade)
	p.setCellValue(f, "Сводка", summaryRow, 'M'-'A', data.AvgGrade)

	// Записываем общую сумму (столбец L)
	p.setCellValue(f, sheetName, courseRow, 'L'-'A', data.TotalSum)
	p.setCellValue(f, "Сводка", summaryRow, 'L'-'A', data.TotalSum)

	// Записываем финальную оценку (столбец O)
	p.setCellValue(f, sheetName, courseRow, 'O'-'A', data.FinalGrade)
	p.setCellValue(f, "Сводка", summaryRow, 'O'-'A', data.FinalGrade)

	// Записываем название группы (столбец G)
	groupName := strings.TrimSuffix(filePath, ".xls")
	p.setCellValue(f, sheetName, courseRow, 'G'-'A', groupName)
	p.setCellValue(f, "Сводка", summaryRow, 'G'-'A', groupName)
}

// GetGroupNumber извлекает номер группы из имени файла
func (p *Processor) GetGroupNumber(filePath string) (string, error) {
	index := strings.Index(filePath, ".")
	if index == -1 || len(filePath) <= index+2 {
		return "", fmt.Errorf("некорректный формат имени файла: %s", filePath)
	}

	return filePath[index+1 : index+3], nil
}
