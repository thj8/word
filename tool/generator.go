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

var EnglishColumnStyleTop = []excelize.Border{
	{Type: "left", Color: "000000", Style: 1},
	{Type: "right", Color: "000000", Style: 1},
	{Type: "top", Color: "000000", Style: 1},
	{Type: "bottom", Color: "000000", Style: 3}, // 底部虚线
}

var EnglishColumnStyleMiddle = []excelize.Border{
	{Type: "left", Color: "000000", Style: 1},
	{Type: "right", Color: "000000", Style: 1},
	{Type: "top", Color: "000000", Style: 3},    // 顶部虚线
	{Type: "bottom", Color: "000000", Style: 3}, // 底部虚线
}

var EnglishColumnStyleBottom = []excelize.Border{
	{Type: "left", Color: "000000", Style: 1},
	{Type: "right", Color: "000000", Style: 1},
	{Type: "top", Color: "000000", Style: 3}, // 顶部虚线
	{Type: "bottom", Color: "000000", Style: 1},
}

// generateWordExercise 生成单词默写练习表的内部函数
func GenerateWordExerciseWithStartIndex(f *excelize.File, sheet string, words []string, startIndex int) error {
	if len(words) == 0 {
		return fmt.Errorf("单词列表不能为空")
	}

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
	englishColumnStyle_top, _ := f.NewStyle(&excelize.Style{
		Border: EnglishColumnStyleTop,
	})
	englishColumnStyle_middle, _ := f.NewStyle(&excelize.Style{
		Border: EnglishColumnStyleMiddle,
	})
	englishColumnStyle_bottom, _ := f.NewStyle(&excelize.Style{
		Border: EnglishColumnStyleBottom,
	})

	// =====================
	// 单词内容 - 动态分配左右两侧
	// =====================
	totalWords := len(words)
	leftCount := (totalWords + 1) / 2    // 左侧数量（如果奇数则左侧多一个）
	rightCount := totalWords - leftCount // 右侧数量

	var leftWords, rightWords []string
	if totalWords > 0 {
		leftWords = words[0:leftCount]
		rightWords = words[leftCount:]
	}

	// 计算需要多少行（取决于较多的那一侧）
	maxCount := leftCount
	if rightCount > maxCount {
		maxCount = rightCount
	}

	startRow := 4
	for i := 0; i < maxCount; i++ {
		r := startRow + i*3 // 每个单词条目占据3行

		// 左侧 - 如果还有单词
		if i < leftCount {
			f.SetCellValue(sheet, fmt.Sprintf("A%d", r), i+1+startIndex)
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
			f.SetCellStyle(sheet, fmt.Sprintf("C%d", r), fmt.Sprintf("C%d", r), englishColumnStyle_top)
			f.SetCellStyle(sheet, fmt.Sprintf("C%d", r+1), fmt.Sprintf("C%d", r+1), englishColumnStyle_middle)
			f.SetCellStyle(sheet, fmt.Sprintf("C%d", r+2), fmt.Sprintf("C%d", r+2), englishColumnStyle_bottom)
		}

		// 右侧 - 如果还有单词
		if i < rightCount {
			f.SetCellValue(sheet, fmt.Sprintf("D%d", r), i+leftCount+1+startIndex)
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
			f.SetCellStyle(sheet, fmt.Sprintf("F%d", r), fmt.Sprintf("F%d", r), englishColumnStyle_top)
			f.SetCellStyle(sheet, fmt.Sprintf("F%d", r+1), fmt.Sprintf("F%d", r+1), englishColumnStyle_middle)
			f.SetCellStyle(sheet, fmt.Sprintf("F%d", r+2), fmt.Sprintf("F%d", r+2), englishColumnStyle_bottom)
		}
	}

	return nil
}

// GenerateWordExerciseTable 生成单词默写练习表的便捷函数
// 参数: words - 包含中文词语的切片
// 参数: filename - 输出Excel文件的名称
func GenerateWordExerciseTable(words []string, filename string) error {
	f := excelize.NewFile()

	// 获取默认工作表的名称并重命名为所需名称
	sheetName := "Sheet1"
	activeSheetIndex := f.GetActiveSheetIndex()
	f.SetSheetName(f.GetSheetName(activeSheetIndex), sheetName)

	err := GenerateWordExerciseWithStartIndex(f, sheetName, words, 0)
	if err != nil {
		return err
	}

	return f.SaveAs(filename)
}

// GenExerSheetWithNames 生成多个工作表的单词默写练习表
// 参数: datasets - 包含多个单词列表的切片
// 参数: sheetNames - 包含工作表名称的切片
// 参数: filename - 输出Excel文件的名称
func GenExerSheetWithNames(datasets [][]string, sheetNames []string, filename string) error {
	if len(datasets) != len(sheetNames) {
		return fmt.Errorf("数据集数量与工作表名称数量不匹配")
	}

	f := excelize.NewFile()

	// 遍历所有数据集并创建工作表
	currentIndex := 0
	for i, words := range datasets {
		sheetName := sheetNames[i]

		// 对于第一个工作表，重命名默认工作表而不是创建新的
		if i == 0 {
			activeSheetIndex := f.GetActiveSheetIndex()
			f.SetSheetName(f.GetSheetName(activeSheetIndex), sheetName)
		} else {
			// 为后续工作表创建新工作表
			f.NewSheet(sheetName)
		}

		if len(words) == 0 {
			continue // 跳过空数据集
		}

		// 使用当前索引作为起始序号
		err := GenerateWordExerciseWithStartIndex(f, sheetName, words, currentIndex)
		if err != nil {
			return err
		}

		// 更新下一个工作表的起始序号
		currentIndex += len(words)
	}

	return f.SaveAs(filename)
}

// GenExerSheet 生成多个工作表的单词默写练习表（使用默认页面名称）
// 参数: allWords - 包含所有单词的切片
// 参数: filename - 输出Excel文件的名称
// 每个sheet最多包含40个单词，超出的部分会自动创建新sheet
func GenExerSheet(allWords []string, filename string) error {
	// 每个sheet最多包含40个单词
	const maxWordsPerSheet = 40
	
	// 计算需要多少个工作表
	numSheets := len(allWords) / maxWordsPerSheet
	if len(allWords)%maxWordsPerSheet > 0 {
		numSheets++
	}
	
	// 如果没有单词，至少创建一个空的工作表
	if numSheets == 0 {
		numSheets = 1
	}
	
	// 将单词列表分割成多个数据集
	datasets := make([][]string, numSheets)
	for i := 0; i < numSheets; i++ {
		start := i * maxWordsPerSheet
		end := start + maxWordsPerSheet
		if end > len(allWords) {
			end = len(allWords)
		}
		datasets[i] = allWords[start:end]
	}
	
	// 生成默认的工作表名称 (Page1, Page2, ...)
	sheetNames := make([]string, numSheets)
	for i := range datasets {
		sheetNames[i] = fmt.Sprintf("Page%d", i+1)
	}

	return GenExerSheetWithNames(datasets, sheetNames, filename)
}
