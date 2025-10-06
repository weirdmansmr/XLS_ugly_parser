package main

import (
	"XLS_ugly_parser/config"
	"XLS_ugly_parser/exporter"
	"XLS_ugly_parser/processor"
	"XLS_ugly_parser/sorter"
	"log"

	"github.com/xuri/excelize/v2"
)

func main() {
	// Инициализируем компоненты
	proc := processor.NewProcessor()
	sort := sorter.NewSorter()
	exp := exporter.NewExporter()

	// Создаем Excel файл
	f := excelize.NewFile()

	// Создаем листы
	for _, sheetName := range config.SheetNames {
		f.NewSheet(sheetName)
	}

	// Настраиваем отложенное сохранение результата
	defer func() {
		if err := f.SaveAs("result.xlsx"); err != nil {
			log.Fatalf("Ошибка сохранения файла: %v", err)
		}
	}()

	// Инициализируем карту начальных строк для каждого листа
	startRowMap := initializeStartRowMap()

	// Обрабатываем все файлы
	processAllFiles(proc, f, startRowMap)

	// Сортируем все листы
	sortAllSheets(sort, f, startRowMap)

	// Экспортируем топ-10 в JavaScript файл
	if err := exp.ExportTop10ToFile(f, "top10.js"); err != nil {
		log.Printf("Ошибка экспорта топ-10: %v", err)
	}
}

// initializeStartRowMap создает карту начальных строк для каждого листа
func initializeStartRowMap() map[string]int {
	startRowMap := make(map[string]int)
	for _, sheetName := range config.SheetNames {
		startRowMap[sheetName] = 2 // Начинаем со 2-й строки (1-я - заголовок)
	}
	return startRowMap
}

// processAllFiles обрабатывает все файлы из конфигурации
func processAllFiles(proc *processor.Processor, f *excelize.File, startRowMap map[string]int) {
	for _, filePath := range config.Files {
		// Извлекаем номер группы из имени файла
		groupNumber, err := proc.GetGroupNumber(filePath)
		if err != nil {
			log.Printf("Ошибка при обработке имени файла: %v", err)
			continue
		}

		// Определяем название листа по номеру группы
		sheetName, ok := config.CourseMapping[groupNumber]
		if !ok {
			log.Printf("Пропускаем файл с номером группы %s: %s", groupNumber, filePath)
			continue
		}

		// Обрабатываем файл
		startRowMap[sheetName] = proc.ProcessFile(filePath, f, startRowMap[sheetName], sheetName, startRowMap)
	}
}

// sortAllSheets сортирует все листы по финальной оценке
func sortAllSheets(sort *sorter.Sorter, f *excelize.File, startRowMap map[string]int) {
	for _, sheet := range config.SheetNames {
		sort.SortSheetByColumn(f, sheet, startRowMap[sheet])
	}
}
