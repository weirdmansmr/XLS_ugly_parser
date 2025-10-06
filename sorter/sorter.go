package sorter

import (
	"fmt"
	"log"
	"strconv"

	"github.com/xuri/excelize/v2"
)

// Sorter отвечает за сортировку данных в Excel
type Sorter struct{}

// NewSorter создает новый экземпляр сортировщика
func NewSorter() *Sorter {
	return &Sorter{}
}

// SortSheetByColumn сортирует лист по столбцу O (финальная оценка) по убыванию
func (s *Sorter) SortSheetByColumn(f *excelize.File, sheetName string, lastRow int) {
	if lastRow <= 2 {
		return // Нет данных для сортировки
	}

	dataRange := fmt.Sprintf("A2:O%d", lastRow-1)

	// Устанавливаем автофильтр
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

	// Получаем все строки
	rows, err := f.GetRows(sheetName)
	if err != nil {
		log.Printf("Ошибка получения строк для %s: %v", sheetName, err)
		return
	}

	if len(rows) <= 2 {
		return
	}

	// Пропускаем заголовок
	dataRows := rows[1:]

	// Сортируем по столбцу O (индекс 14) по убыванию
	s.bubbleSortByColumn(dataRows, 14)

	// Записываем отсортированные данные обратно
	s.writeRowsToSheet(f, sheetName, dataRows)
}

// bubbleSortByColumn сортирует строки пузырьковым методом по указанному столбцу
func (s *Sorter) bubbleSortByColumn(rows [][]string, columnIndex int) {
	for i := 0; i < len(rows)-1; i++ {
		for j := 0; j < len(rows)-i-1; j++ {
			val1 := s.getFloatValue(rows[j], columnIndex)
			val2 := s.getFloatValue(rows[j+1], columnIndex)

			// Сортируем по убыванию (большие значения первыми)
			if val1 < val2 {
				rows[j], rows[j+1] = rows[j+1], rows[j]
			}
		}
	}
}

// getFloatValue извлекает float значение из ячейки
func (s *Sorter) getFloatValue(row []string, columnIndex int) float64 {
	if len(row) > columnIndex {
		if val, err := strconv.ParseFloat(row[columnIndex], 64); err == nil {
			return val
		}
	}
	return 0
}

// writeRowsToSheet записывает строки обратно в лист
func (s *Sorter) writeRowsToSheet(f *excelize.File, sheetName string, rows [][]string) {
	for rowIdx, row := range rows {
		for colIdx, cellValue := range row {
			cell := fmt.Sprintf("%s%d", string('A'+colIdx), rowIdx+2)
			f.SetCellValue(sheetName, cell, cellValue)
		}
	}
}
