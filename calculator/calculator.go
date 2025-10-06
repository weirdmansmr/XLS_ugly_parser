package calculator

import (
	"XLS_ugly_parser/config"
)

// RowData содержит данные одной строки для обработки
type RowData struct {
	Sums          [4]float64
	Values        [4]float64
	AvgGrade      float64
	TotalSum      float64
	FinalGrade    float64
	OnlyDFilled   bool
	InIgnoreRange bool
}

// Calculator содержит логику вычислений
type Calculator struct{}

// NewCalculator создает новый экземпляр калькулятора
func NewCalculator() *Calculator {
	return &Calculator{}
}

// CalculateRowData вычисляет все необходимые метрики для строки
func (c *Calculator) CalculateRowData(values [4]float64) *RowData {
	data := &RowData{
		Values: values,
	}

	// Вычисляем суммы с коэффициентами
	for i, val := range values {
		data.Sums[i] = val * config.SumCoefficients[i]
	}

	// Вычисляем средневзвешенную оценку
	data.AvgGrade = c.calculateAvgGrade(values)

	// Вычисляем общую сумму
	data.TotalSum = data.Sums[0] + data.Sums[1] - data.Sums[2] - data.Sums[3]

	// Проверяем условия игнорирования
	data.OnlyDFilled = c.isOnlyDFilled(values)
	data.InIgnoreRange = c.isInIgnoreRange(data.AvgGrade)

	// Вычисляем финальную оценку
	data.FinalGrade = c.calculateFinalGrade(data)

	return data
}

// calculateAvgGrade вычисляет средневзвешенную оценку
func (c *Calculator) calculateAvgGrade(values [4]float64) float64 {
	weightedSum := 0.0
	totalValues := 0.0

	for i, val := range values {
		weightedSum += val * config.WeightCoefficients[i]
		totalValues += val
	}

	if totalValues > 0 {
		return weightedSum / totalValues
	}

	return 0
}

// calculateFinalGrade вычисляет финальную оценку с округлением
func (c *Calculator) calculateFinalGrade(data *RowData) float64 {
	// Если выполняются условия игнорирования, возвращаем 0
	if data.OnlyDFilled || data.InIgnoreRange {
		return 0
	}

	// Вычисляем финальную оценку
	if data.AvgGrade > 0 && (config.FinalGradeDivisor-data.AvgGrade) > 0 {
		finalGrade := data.TotalSum / (config.FinalGradeDivisor - data.AvgGrade)
		// Округляем до 2 знаков после запятой
		return c.roundToTwoDecimals(finalGrade)
	}

	return 0
}

// isOnlyDFilled проверяет, заполнен ли только столбец D
func (c *Calculator) isOnlyDFilled(values [4]float64) bool {
	return values[1] > 0 && values[0] == 0 && values[2] == 0 && values[3] == 0
}

// isInIgnoreRange проверяет, находится ли avgGrade в диапазоне игнорирования
func (c *Calculator) isInIgnoreRange(avgGrade float64) bool {
	return avgGrade >= config.IgnoreRangeMin && avgGrade <= config.IgnoreRangeMax
}

// roundToTwoDecimals округляет число до 2 знаков после запятой
func (c *Calculator) roundToTwoDecimals(value float64) float64 {
	return float64(int(value*100+0.5)) / 100
}
