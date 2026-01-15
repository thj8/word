package tool

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

var FullBorder = []excelize.Border{
	{Type: "left", Color: "000000", Style: 1},
	{Type: "right", Color: "000000", Style: 1},
	{Type: "top", Color: "000000", Style: 1},
	{Type: "bottom", Color: "000000", Style: 1},
}

// generateWordExercise 生成单词默写练习表的内部函数
func GenerateWordExercise(words []string, filename string) error {
	if len(words) != 40 {
		return fmt.Errorf("必须提供40个单词，但提供了 %d 个", len(words))
	}

	f := excelize.NewFile()
	sheet := "Sheet1"

	// 设置页面尺寸
	generator := NewExcelGenerator()
	generator.SetPageSize(f, sheet)

	// =====================
	// 列宽 - 缩小到80%
	// =====================
	f.SetColWidth(sheet, "A", "A", 4.8)  // 6*0.8
	f.SetColWidth(sheet, "B", "B", 16.0) // 20*0.8
	f.SetColWidth(sheet, "C", "C", 16.0) // 20*0.8

	f.SetColWidth(sheet, "D", "D", 4.8)  // 6*0.8
	f.SetColWidth(sheet, "E", "E", 16.0) // 20*0.8
	f.SetColWidth(sheet, "F", "F", 16.0) // 20*0.8

	// =====================
	// 标题
	// =====================
	titleStyle, _ := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{Bold: true, Size: 14},
	})

	f.MergeCell(sheet, "A1", "C1")
	f.SetCellValue(sheet, "A1", "单词默写练习")
	f.SetCellStyle(sheet, "A1", "A1", titleStyle)

	f.MergeCell(sheet, "D1", "F1")
	f.SetCellValue(sheet, "D1", "姓名___________  用时____________")

	// =====================
	// 表头
	// =====================
	headerStyle, _ := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{Bold: true},
		Alignment: &excelize.Alignment{
			Horizontal: "center",
			Vertical:   "center",
		},
		Border: FullBorder,
	})

	headers := []string{"序号", "中文", "英文", "序号", "中文", "英文"}
	for i, h := range headers {
		cell := fmt.Sprintf("%c3", 'A'+i)
		f.SetCellValue(sheet, cell, h)
		f.SetCellStyle(sheet, cell, cell, headerStyle)
	}

	// =====================
	// 样式
	// =====================
	cellStyle, _ := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			Horizontal: "center",
			Vertical:   "center",
		},
		Border: FullBorder,
	})

	// C列和F列特殊边框样式（全边框，但第一个单元格下方线和第二个格子上面线为虚线）
	englishColumnStyle_top1, _ := f.NewStyle(&excelize.Style{
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 3}, // 底部虚线
		},
	})
	englishColumnStyle_middle, _ := f.NewStyle(&excelize.Style{
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 3},    // 顶部虚线
			{Type: "bottom", Color: "000000", Style: 3}, // 底部虚线
		},
	})
	englishColumnStyle_bottom, _ := f.NewStyle(&excelize.Style{
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 3}, // 顶部虚线
			{Type: "bottom", Color: "000000", Style: 1},
		},
	})

	// =====================
	// 单词内容
	// =====================
	leftWords := words[:20]  // 前20个单词
	rightWords := words[20:] // 后20个单词

	startRow := 4
	for i := 0; i < 20; i++ {
		r := startRow + i*3 // 每个单词条目占据3行

		// 左侧
		f.SetCellValue(sheet, fmt.Sprintf("A%d", r), i+1)
		f.SetCellValue(sheet, fmt.Sprintf("B%d", r), leftWords[i])
		f.SetCellStyle(sheet, fmt.Sprintf("A%d", r), fmt.Sprintf("B%d", r), cellStyle)

		// 设置三行高度为10
		f.SetRowHeight(sheet, r, 10)
		f.SetRowHeight(sheet, r+1, 10)
		f.SetRowHeight(sheet, r+2, 10)

		// 合并下方两行
		f.MergeCell(sheet, fmt.Sprintf("A%d", r), fmt.Sprintf("A%d", r+2))
		f.MergeCell(sheet, fmt.Sprintf("B%d", r), fmt.Sprintf("B%d", r+2))

		// 应用样式到合并后的单元格
		f.SetCellStyle(sheet, fmt.Sprintf("A%d", r), fmt.Sprintf("B%d", r+2), cellStyle)
		// 为C列的每个单元格单独设置样式（不合并C列）
		f.SetCellStyle(sheet, fmt.Sprintf("C%d", r), fmt.Sprintf("C%d", r), englishColumnStyle_top1)
		f.SetCellStyle(sheet, fmt.Sprintf("C%d", r+1), fmt.Sprintf("C%d", r+1), englishColumnStyle_middle)
		f.SetCellStyle(sheet, fmt.Sprintf("C%d", r+2), fmt.Sprintf("C%d", r+2), englishColumnStyle_bottom)

		// 右侧
		f.SetCellValue(sheet, fmt.Sprintf("D%d", r), i+21)
		f.SetCellValue(sheet, fmt.Sprintf("E%d", r), rightWords[i])
		f.SetCellStyle(sheet, fmt.Sprintf("D%d", r), fmt.Sprintf("E%d", r), cellStyle)

		// 设置三行高度为10
		f.SetRowHeight(sheet, r, 10)
		f.SetRowHeight(sheet, r+1, 10)
		f.SetRowHeight(sheet, r+2, 10)

		// 合并下方两行
		f.MergeCell(sheet, fmt.Sprintf("D%d", r), fmt.Sprintf("D%d", r+2))
		f.MergeCell(sheet, fmt.Sprintf("E%d", r), fmt.Sprintf("E%d", r+2))

		// 应用样式到合并后的单元格
		f.SetCellStyle(sheet, fmt.Sprintf("D%d", r), fmt.Sprintf("E%d", r+2), cellStyle)
		// 为F列的每个单元格单独设置样式（不合并F列）
		f.SetCellStyle(sheet, fmt.Sprintf("F%d", r), fmt.Sprintf("F%d", r), englishColumnStyle_top1)
		f.SetCellStyle(sheet, fmt.Sprintf("F%d", r+1), fmt.Sprintf("F%d", r+1), englishColumnStyle_middle)
		f.SetCellStyle(sheet, fmt.Sprintf("F%d", r+2), fmt.Sprintf("F%d", r+2), englishColumnStyle_bottom)
	}

	// =====================
	// 保存
	// =====================
	if err := f.SaveAs(filename); err != nil {
		return err
	}

	return nil
}

// GenerateWordExerciseTable 生成单词默写练习表的便捷函数
// 参数: words - 包含40个中文词语的切片
// 参数: filename - 输出Excel文件的名称
func GenerateWordExerciseTable(words []string, filename string) error {
	return GenerateWordExercise(words, filename)
}
