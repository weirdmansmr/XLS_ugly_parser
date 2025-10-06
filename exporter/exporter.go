package exporter

import (
	"XLS_ugly_parser/config"
	"XLS_ugly_parser/models"
	"fmt"
	"log"
	"os"
	"strconv"

	"github.com/xuri/excelize/v2"
)

// Exporter отвечает за экспорт данных в JavaScript формат
type Exporter struct{}

// NewExporter создает новый экземпляр экспортера
func NewExporter() *Exporter {
	return &Exporter{}
}

// ExportTop10ToFile экспортирует топ-10 студентов всех курсов в JS файл
func (e *Exporter) ExportTop10ToFile(f *excelize.File, outputPath string) error {
	jsOutput := "// Топ-10 студентов по курсам\n\n"

	for _, sheet := range config.SheetNames {
		top10 := e.getTop10Students(f, sheet)
		jsOutput += e.formatAsJS(top10, config.CourseJSNames[sheet])
		jsOutput += "\n"
	}

	// Сохраняем результат в файл
	err := os.WriteFile(outputPath, []byte(jsOutput), 0644)
	if err != nil {
		return fmt.Errorf("ошибка сохранения файла %s: %w", outputPath, err)
	}

	log.Printf("Файл %s успешно создан", outputPath)
	return nil
}

// getTop10Students извлекает топ-10 студентов из листа
func (e *Exporter) getTop10Students(f *excelize.File, sheetName string) []models.Student {
	rows, err := f.GetRows(sheetName)
	if err != nil || len(rows) <= 1 {
		return []models.Student{}
	}

	students := []models.Student{}
	maxRows := 11 // 1 заголовок + 10 строк данных

	if len(rows) < maxRows {
		maxRows = len(rows)
	}

	for i := 1; i < maxRows; i++ {
		row := rows[i]
		if len(row) < 15 {
			continue
		}

		student := e.parseStudentFromRow(row)
		
		// Пропускаем строки с GPA = 0 или пустыми данными
		if student.GPA > 0 && student.Name != "" && student.Group != "" {
			student.ID = len(students)
			students = append(students, student)
		}

		// Если уже собрали 10 студентов, останавливаемся
		if len(students) >= 10 {
			break
		}
	}

	return students
}

// parseStudentFromRow парсит данные студента из строки
func (e *Exporter) parseStudentFromRow(row []string) models.Student {
	student := models.Student{}

	// Имя (столбец B, индекс 1)
	if len(row) > 1 {
		student.Name = row[1]
	}

	// Группа (столбец G, индекс 6)
	if len(row) > 6 {
		student.Group = row[6]
	}

	// GPA (столбец O, индекс 14)
	if len(row) > 14 {
		if val, err := strconv.ParseFloat(row[14], 64); err == nil {
			student.GPA = val
		}
	}

	return student
}

// formatAsJS форматирует список студентов в JavaScript массив
func (e *Exporter) formatAsJS(students []models.Student, courseName string) string {
	result := fmt.Sprintf("const %s = [\n", courseName)

	for _, student := range students {
		result += "  {\n"
		result += fmt.Sprintf("    id: %d,\n", student.ID)
		result += fmt.Sprintf("    group: \"%s\",\n", student.Group)
		result += fmt.Sprintf("    name: \"%s\",\n", student.Name)
		result += fmt.Sprintf("    gpa: %.2f,\n", student.GPA)
		result += "  },\n"
	}

	result += "];\n"
	return result
}
