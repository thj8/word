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

// CreateTitleStyle 创建标题样式
func (eg *ExcelGenerator) CreateTitleStyle(f *excelize.File) (int, error) {
	style := &excelize.Style{
		Font: &excelize.Font{
			Bold: false,
			Size: 12,
		},
		Alignment: &excelize.Alignment{
			Horizontal: "center",
			Vertical:   "center",
		},
	}
	return f.NewStyle(style)
}

// CreateDateStyle 创建日期样式
func (eg *ExcelGenerator) CreateDateStyle(f *excelize.File) (int, error) {
	style := &excelize.Style{
		Font: &excelize.Font{
			Bold: true,
			Size: 12,
		},
		Alignment: &excelize.Alignment{
			Horizontal: "left",
			Vertical:   "center",
		},
	}
	return f.NewStyle(style)
}

// CreateTableHeaderStyle 创建表头样式
func (eg *ExcelGenerator) CreateTableHeaderStyle(f *excelize.File) (int, error) {
	style := &excelize.Style{
		Font: &excelize.Font{
			Bold: true,
			Size: 12,
		},
		Alignment: &excelize.Alignment{
			Horizontal: "center",
			Vertical:   "center",
		},
		Fill: excelize.Fill{
			Type:    "pattern",
			Pattern: 1,
			Color:   []string{"#E0E0E0"},
		},
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
		},
	}
	return f.NewStyle(style)
}

// CreateCellStyle 创建单元格样式
func (eg *ExcelGenerator) CreateCellStyle(f *excelize.File) (int, error) {
	style := &excelize.Style{
		Alignment: &excelize.Alignment{
			Horizontal: "center",
			Vertical:   "center",
		},
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
		},
	}
	return f.NewStyle(style)
}

// CreateEnglishColumnStyleTop 创建英文列顶部样式
func (eg *ExcelGenerator) CreateEnglishColumnStyleTop(f *excelize.File) (int, error) {
	style := &excelize.Style{
		Alignment: &excelize.Alignment{
			Horizontal: "center",
			Vertical:   "center",
		},
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 3}, // 底部虚线
		},
	}
	return f.NewStyle(style)
}

// CreateEnglishColumnStyleMiddle 创建英文列中间样式
func (eg *ExcelGenerator) CreateEnglishColumnStyleMiddle(f *excelize.File) (int, error) {
	style := &excelize.Style{
		Alignment: &excelize.Alignment{
			Horizontal: "center",
			Vertical:   "center",
		},
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 3},    // 顶部虚线
			{Type: "bottom", Color: "000000", Style: 3}, // 底部虚线
		},
	}
	return f.NewStyle(style)
}

// CreateEnglishColumnStyleBottom 创建英文列底部样式
func (eg *ExcelGenerator) CreateEnglishColumnStyleBottom(f *excelize.File) (int, error) {
	style := &excelize.Style{
		Alignment: &excelize.Alignment{
			Horizontal: "center",
			Vertical:   "center",
		},
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 3}, // 顶部虚线
			{Type: "bottom", Color: "000000", Style: 1},
		},
	}
	return f.NewStyle(style)
}