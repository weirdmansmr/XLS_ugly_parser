package models

// Student представляет данные студента для экспорта
type Student struct {
	ID    int     `json:"id"`
	Group string  `json:"group"`
	Name  string  `json:"name"`
	GPA   float64 `json:"gpa"`
}
