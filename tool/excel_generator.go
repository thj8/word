package tool

import (
	"github.com/xuri/excelize/v2"
)

// ExcelGenerator 是一个Excel生成器接口
type ExcelGenerator struct{}

// NewExcelGenerator 创建一个新的Excel生成器
func NewExcelGenerator() *ExcelGenerator {
	return &ExcelGenerator{}
}

// GenerateWordExercise 生成单词默写练习表
func (eg *ExcelGenerator) GenerateWordExercise(words []string, filename string) error {
	return GenerateWordExercise(words, filename)
}

// SetPageSize 设置页面尺寸
func (eg *ExcelGenerator) SetPageSize(f *excelize.File, sheet string) {
	// 设置为A4纸张大小
	orientation := "portrait"
	size := 9 // A4纸张
	f.SetPageLayout(sheet, &excelize.PageLayoutOptions{
		Orientation: &orientation,
		Size:        &size,
	})

	// 设置页边距
	leftMargin := 0.7
	rightMargin := 0.7
	topMargin := 0.75
	bottomMargin := 0.75
	f.SetPageMargins(sheet, &excelize.PageLayoutMarginsOptions{
		Left:   &leftMargin,
		Right:  &rightMargin,
		Top:    &topMargin,
		Bottom: &bottomMargin,
	})
}
