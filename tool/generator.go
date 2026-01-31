package tool

import (
	"fmt"
	"math/rand"
	"path/filepath"
	"strings"
	"time"

	"github.com/thj8/word/utils"

	"github.com/xuri/excelize/v2"
)

// ExerciseGenerator 练习表生成器结构体
type ExerciseGenerator struct {
	Opts            GenerateOptions // 生成选项
	ResourceName    string          // 资源名称
	OriginalWords   []string        // 原始单词数据
	ProcessedWords  []string        // 处理后的单词数据（截取、打乱等）
	SplitWordGroups [][]string      // 按每页40个单词分割的单词组
}

// GenerateOptions 定义生成Excel文件的选项
type GenerateOptions struct {
	ShowPos   bool // 是否显示词性
	WordCount int  // 输出单词个数，-1表示全部
	Shuffle   bool // 是否随机乱序
}

// NewExerciseGenerator 创建新的练习表生成器
func NewExerciseGenerator(resourceName string, opts GenerateOptions, originalWords []string) *ExerciseGenerator {
	generator := &ExerciseGenerator{
		ResourceName:  resourceName,
		Opts:          opts,
		OriginalWords: originalWords,
	}
	generator.ProcessWords()         // 初始化时处理单词
	generator.SplitWordsIntoGroups() // 初始化时分割单词
	return generator
}

// genWordExerciseInternal 内部生成函数
func (eg *ExerciseGenerator) genWordExerciseInternal(f *excelize.File, sheet string, words []string, startIndex int) error {
	if len(words) == 0 {
		return fmt.Errorf("单词列表不能为空")
	}

	// 设置页面尺寸
	generator := NewExcelGenerator()
	generator.SetPageSize(f, sheet)

	// 创建样式
	titleStyle, _ := generator.TitleStyle(f)
	tableHeaderStyle, _ := generator.TblHdrStyle(f)
	cellStyle, _ := generator.CellStyle(f)
	englishColumnStyle_top, _ := generator.EngColStyTop(f)
	englishColumnStyle_middle, _ := generator.EngColStyMid(f)
	englishColumnStyle_bottom, _ := generator.EngColStyBot(f)

	// 写入标题
	title := fmt.Sprintf("%s - %s - 默写", eg.ResourceName, sheet)
	f.SetCellValue(sheet, "A1", title)
	f.MergeCell(sheet, "A1", "C1")
	f.SetCellStyle(sheet, "A1", "C1", titleStyle)

	// 写入姓名和日期区域
	f.SetCellValue(sheet, "D1", "姓名___________  用时____________")
	f.MergeCell(sheet, "D1", "F1")
	f.SetCellStyle(sheet, "D1", "F1", titleStyle)

	// 设置列宽
	f.SetColWidth(sheet, "A", "A", 4.86) // 6*0.81
	f.SetColWidth(sheet, "B", "B", 16.2) // 20*0.81
	f.SetColWidth(sheet, "C", "C", 16.2) // 20*0.81
	f.SetColWidth(sheet, "D", "D", 4.86) // 6*0.81
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
	if len(words)%2 == 0 {
		maxCount = len(words) / 2
	} else {
		maxCount = len(words)/2 + 1
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

			// 根据Opts选项决定如何显示左侧单词
			displayText := leftWords[i]
			if !eg.Opts.ShowPos {
				// 如果不显示pos，则尝试提取纯文本部分
				displayText = utils.ExtractTextWithoutPos(leftWords[i])
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

			// 根据Opts选项决定如何显示右侧单词
			displayText := rightWords[i]
			if !eg.Opts.ShowPos {
				// 如果不显示pos，则尝试提取纯文本部分
				displayText = utils.ExtractTextWithoutPos(rightWords[i])
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

// ProcessWords 根据选项处理原始单词列表（截取、打乱等）
func (eg *ExerciseGenerator) ProcessWords() {
	if len(eg.OriginalWords) == 0 {
		eg.ProcessedWords = []string{}
		return
	}

	processedWords := make([]string, len(eg.OriginalWords))
	copy(processedWords, eg.OriginalWords)

	// 处理shuffle参数
	if eg.Opts.Shuffle {
		// 如果需要打乱，使用随机种子打乱单词顺序
		rand.Seed(time.Now().UnixNano())
		for i := len(processedWords) - 1; i > 0; i-- {
			j := rand.Intn(i + 1)
			processedWords[i], processedWords[j] = processedWords[j], processedWords[i]
		}
	}

	// 处理wordCount参数
	if eg.Opts.WordCount > 0 && eg.Opts.WordCount < len(processedWords) {
		processedWords = processedWords[:eg.Opts.WordCount]
	}

	eg.ProcessedWords = processedWords
}

// SplitWordsIntoGroups 按每页40个单词分割单词列表
func (eg *ExerciseGenerator) SplitWordsIntoGroups() {
	pageSize := 40
	if len(eg.ProcessedWords) == 0 {
		eg.SplitWordGroups = [][]string{}
		return
	}

	numSheets := len(eg.ProcessedWords) / pageSize
	if len(eg.ProcessedWords)%pageSize > 0 {
		numSheets++
	}

	eg.SplitWordGroups = make([][]string, numSheets)
	for i := 0; i < numSheets; i++ {
		start := i * pageSize
		end := start + pageSize
		if end > len(eg.ProcessedWords) {
			end = len(eg.ProcessedWords)
		}
		eg.SplitWordGroups[i] = eg.ProcessedWords[start:end]
	}
}

// genMultiSheetExercise 生成多个工作表的单词默写练习表，支持显示pos选项(内部函数，私有)
func (eg *ExerciseGenerator) genMultiSheetExercise(sheetNames []string, filename string) error {
	f := excelize.NewFile()

	// 遍历所有分割好的单词组并创建工作表
	currentIndex := 0
	for i, words := range eg.SplitWordGroups {
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

		err := eg.genWordExerciseInternal(f, sheetName, words, currentIndex)
		if err != nil {
			return err
		}

		// 更新下一个工作表的起始序号
		currentIndex += len(words)
	}

	return f.SaveAs(filename)
}

// GenerateFilename 根据资源名称和选项生成Excel文件名
func (eg *ExerciseGenerator) GenerateFilename() string {
	return GenerateExcelFilename(eg.ResourceName, eg.Opts.ShowPos, eg.Opts.WordCount, eg.Opts.Shuffle)
}

// Generate 生成练习表
func (eg *ExerciseGenerator) Generate(filename string) error {
	if len(eg.OriginalWords) == 0 {
		return fmt.Errorf("单词列表不能为空")
	}

	// 如果没有提供文件名，则自动生成
	if filename == "" {
		filename = eg.GenerateFilename()
	}

	// 生成工作表名称
	sheetNames := make([]string, len(eg.SplitWordGroups))
	for i := 0; i < len(eg.SplitWordGroups); i++ {
		if len(eg.SplitWordGroups) == 1 {
			sheetNames[i] = "Sheet1"
		} else {
			sheetNames[i] = fmt.Sprintf("Page%d", i+1)
		}
	}

	return eg.genMultiSheetExercise(sheetNames, filename)
}

// GenerateAuto 自动根据资源名称生成练习表
func (eg *ExerciseGenerator) GenerateAuto() error {
	filename := eg.GenerateFilename()
	return eg.Generate(filename)
}

// GenExerciseSheet 生成单词默写练习表（保持向后兼容，默认显示pos）
func GenExerciseSheet(resourceName string, allWords []string, filename string, shuffle bool) error {
	opts := GenerateOptions{
		ShowPos:   true,
		WordCount: -1,
		Shuffle:   shuffle,
	}

	generator := NewExerciseGenerator(resourceName, opts, allWords)
	return generator.Generate(filename)
}

// GenerateExcelFilename 根据资源名称和选项生成Excel文件名
func GenerateExcelFilename(resourceName string, showPos bool, wordCount int, shuffle bool) string {
	cleanResourceName := utils.CleanFileName(resourceName)
	filenameParts := []string{cleanResourceName}
	if !showPos {
		filenameParts = append(filenameParts, "no_pos")
	}
	if wordCount > 0 {
		filenameParts = append(filenameParts, fmt.Sprintf("%dwords", wordCount))
	}
	if shuffle {
		filenameParts = append(filenameParts, "shuffle")
	}
	return filepath.Join("excel", fmt.Sprintf("%s.xlsx", strings.Join(filenameParts, "_")))
}
