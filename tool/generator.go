package tool

import (
	"fmt"
	"math/rand"
	"time"

	"github.com/xuri/excelize/v2"
)

// GenerateOptions 定义生成Excel文件的选项
type GenerateOptions struct {
	ShowPos   bool // 是否显示词性
	WordCount int  // 输出单词个数，-1表示全部
	Shuffle   bool // 是否随机乱序
}

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

// GenWordExerciseInternal 内部生成函数
func GenWordExerciseInternal(f *excelize.File, sheet string, words []string, startIndex int, opts GenerateOptions) error {
	if len(words) == 0 {
		return fmt.Errorf("单词列表不能为空")
	}

	// 设置页面尺寸
	generator := NewExcelGenerator()
	generator.SetPageSize(f, sheet)

	// 创建样式
	titleStyle, _ := generator.CreateTitleStyle(f)
	tableHeaderStyle, _ := generator.CreateTableHeaderStyle(f)
	cellStyle, _ := generator.CreateCellStyle(f)
	englishColumnStyle_top, _ := generator.CreateEnglishColumnStyleTop(f)
	englishColumnStyle_middle, _ := generator.CreateEnglishColumnStyleMiddle(f)
	englishColumnStyle_bottom, _ := generator.CreateEnglishColumnStyleBottom(f)

	// 写入标题
	f.SetCellValue(sheet, "A1", "单词默写练习")
	f.MergeCell(sheet, "A1", "C1")
	f.SetCellStyle(sheet, "A1", "C1", titleStyle)

	// 写入姓名和日期区域
	f.SetCellValue(sheet, "D1", "姓名___________  用时____________")
	f.MergeCell(sheet, "D1", "F1")
	f.SetCellStyle(sheet, "D1", "F1", titleStyle)

	// 设置列宽
	f.SetColWidth(sheet, "A", "A", 4.86)  // 6*0.81
	f.SetColWidth(sheet, "B", "B", 16.2) // 20*0.81
	f.SetColWidth(sheet, "C", "C", 16.2) // 20*0.81
	f.SetColWidth(sheet, "D", "D", 4.86)  // 6*0.81
	f.SetColWidth(sheet, "E", "E", 16.2) // 20*0.81
	f.SetColWidth(sheet, "F", "F", 16.2) // 20*0.81

	// 写入表头
	f.SetCellValue(sheet, "A3", "序号")
	f.SetCellValue(sheet, "B3", "中文")
	f.SetCellValue(sheet, "C3", "英文")
	f.SetCellValue(sheet, "D3", "序号")
	f.SetCellValue(sheet, "E3", "中文")
	f.SetCellValue(sheet, "F3", "英文")
	f.SetCellStyle(sheet, "A3", "F3", tableHeaderStyle)

	// 计算左右两列的单词数
	maxCount := 0
	if len(words) % 2 == 0 {
		maxCount = len(words) / 2
	} else {
		maxCount = len(words) / 2 + 1
	}
	leftCount := maxCount
	rightCount := len(words) - leftCount

	// 确定左右两列的单词列表
	leftWords := words[:leftCount]
	rightWords := words[leftCount:]

	// 确保右侧单词列表长度正确
	if len(rightWords) < rightCount {
		rightCount = len(rightWords)
	}

	// 生成表格内容
	startRow := 4
	for i := 0; i < maxCount; i++ {
		r := startRow + i*3 // 每个单词条目占据3行

		// 左侧 - 如果还有单词
		if i < leftCount {
			f.SetCellValue(sheet, fmt.Sprintf("A%d", r), i+1+startIndex)
			
			// 根据opts选项决定如何显示左侧单词
			displayText := leftWords[i]
			if !opts.ShowPos {
				// 如果不显示pos，则尝试提取纯文本部分
				displayText = extractTextWithoutPos(leftWords[i])
			}
			f.SetCellValue(sheet, fmt.Sprintf("B%d", r), displayText)
			f.SetCellStyle(sheet, fmt.Sprintf("A%d", r), fmt.Sprintf("B%d", r), cellStyle)

			// 设置三行高度为11
			f.SetRowHeight(sheet, r, 11)
			f.SetRowHeight(sheet, r+1, 11)
			f.SetRowHeight(sheet, r+2, 11)

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
			
			// 根据opts选项决定如何显示右侧单词
			displayText := rightWords[i]
			if !opts.ShowPos {
				// 如果不显示pos，则尝试提取纯文本部分
				displayText = extractTextWithoutPos(rightWords[i])
			}
			f.SetCellValue(sheet, fmt.Sprintf("E%d", r), displayText)
			f.SetCellStyle(sheet, fmt.Sprintf("D%d", r), fmt.Sprintf("E%d", r), cellStyle)

			// 设置三行高度为11
			f.SetRowHeight(sheet, r, 11)
			f.SetRowHeight(sheet, r+1, 11)
			f.SetRowHeight(sheet, r+2, 11)

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

// genExerSheetWithNamesAndOptions 生成多个工作表的单词默写练习表，支持显示pos选项(内部函数，私有)
func genExerSheetWithNamesAndOptions(datasets [][]string, sheetNames []string, filename string, showPos bool) error {
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
		opts := GenerateOptions{
			ShowPos: showPos,
		}
		err := GenWordExerciseInternal(f, sheetName, words, currentIndex, opts)
		if err != nil {
			return err
		}

		// 更新下一个工作表的起始序号
		currentIndex += len(words)
	}

	return f.SaveAs(filename)
}

// GenerateExerciseSheet 统一的对外函数，使用Option结构体传参
func GenerateExerciseSheet(allWords []string, filename string, opts GenerateOptions) error {
	if len(allWords) == 0 {
		return fmt.Errorf("单词列表不能为空")
	}

	// 处理shuffle参数
	processedWords := make([]string, len(allWords))
	copy(processedWords, allWords)
	if opts.Shuffle {
		// 如果需要打乱，使用随机种子打乱单词顺序
		rand.Seed(time.Now().UnixNano())
		for i := len(processedWords) - 1; i > 0; i-- {
			j := rand.Intn(i + 1)
			processedWords[i], processedWords[j] = processedWords[j], processedWords[i]
		}
	}

	// 处理wordCount参数
	if opts.WordCount > 0 && opts.WordCount < len(processedWords) {
		processedWords = processedWords[:opts.WordCount]
	}

	// 计算需要多少个工作表（每表最多40个单词）
	pageSize := 40
	numSheets := len(processedWords) / pageSize
	if len(processedWords)%pageSize > 0 {
		numSheets++
	}

	// 分割单词列表
	datasets := make([][]string, numSheets)
	for i := 0; i < numSheets; i++ {
		start := i * pageSize
		end := start + pageSize
		if end > len(processedWords) {
			end = len(processedWords)
		}
		datasets[i] = processedWords[start:end]
	}

	// 生成工作表名称
	sheetNames := make([]string, numSheets)
	for i := 0; i < numSheets; i++ {
		if numSheets == 1 {
			sheetNames[i] = "Sheet1"
		} else {
			sheetNames[i] = fmt.Sprintf("Page%d", i+1)
		}
	}

	return genExerSheetWithNamesAndOptions(datasets, sheetNames, filename, opts.ShowPos)
}

// GenExerciseSheet 生成单词默写练习表（保持向后兼容，默认显示pos）
func GenExerciseSheet(allWords []string, filename string, shuffle bool) error {
	opts := GenerateOptions{
		ShowPos:   true,
		WordCount: -1,
		Shuffle:   shuffle,
	}
	
	return GenerateExerciseSheet(allWords, filename, opts)
}

// extractTextWithoutPos 从单词字符串中提取不包含pos的部分
func extractTextWithoutPos(word string) string {
	// 查找第一个点的位置，通常是pos和text之间的分隔符
	for i, char := range word {
		if char == '.' {
			// 返回点之后的部分（即text部分）
			if i+1 < len(word) {
				return word[i+1:]
			}
			break
		}
	}
	// 如果没有找到点，则返回原字符串
	return word
}