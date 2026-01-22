package tool

import (
	"os"
	"testing"
	"github.com/xuri/excelize/v2"
)

func TestGenExerciseSheet(t *testing.T) {
	// 创建一个临时的单词列表
	words := []string{
		"n.测试单词1",
		"v.测试单词2",
		"adj.测试单词3",
	}

	// 生成临时文件
	tempFile := "temp_test.xlsx"
	
	// 测试正常情况
	err := GenExerciseSheet("test_resource", words, tempFile, false)
	if err != nil {
		t.Errorf("GenExerciseSheet() 函数执行失败: %v", err)
	}
	
	// 检查文件是否创建成功
	if _, err := os.Stat(tempFile); os.IsNotExist(err) {
		t.Error("GenExerciseSheet() 未生成输出文件")
	}
	
	// 清理临时文件
	os.Remove(tempFile)
}

func TestGenExerciseSheetWithShuffle(t *testing.T) {
	// 创建一个临时的单词列表
	words := []string{
		"n.测试单词1",
		"v.测试单词2",
		"adj.测试单词3",
		"adv.测试单词4",
	}

	// 生成临时文件
	tempFile := "temp_test_shuffle.xlsx"
	
	// 测试打乱顺序的情况
	err := GenExerciseSheet("test_resource", words, tempFile, true)
	if err != nil {
		t.Errorf("GenExerciseSheet() with shuffle 函数执行失败: %v", err)
	}
	
	// 检查文件是否创建成功
	if _, err := os.Stat(tempFile); os.IsNotExist(err) {
		t.Error("GenExerciseSheet() with shuffle 未生成输出文件")
	}
	
	// 清理临时文件
	os.Remove(tempFile)
}

func TestGenExerciseSheetEmpty(t *testing.T) {
	// 测试空列表的情况
	words := []string{}

	// 生成临时文件
	tempFile := "temp_test_empty.xlsx"
	
	// 测试空列表的情况
	err := GenExerciseSheet("test_resource", words, tempFile, false)
	if err == nil {
		t.Error("GenExerciseSheet() with empty list 应该返回错误")
	}
	
	// 检查文件是否没有创建
	if _, err := os.Stat(tempFile); err == nil {
		t.Error("GenExerciseSheet() with empty list 不应该生成输出文件")
		os.Remove(tempFile)
	}
}



func TestGenExerciseSheetWithLargeDataset(t *testing.T) {
	// 创建一个较大的数据集，超过40个单词，以测试分页功能
	words := make([]string, 50)
	for i := 0; i < 50; i++ {
		words[i] = "n.测试单词" + string(rune('A'+i%26))
	}

	// 生成临时文件
	tempFile := "temp_test_large.xlsx"
	
	// 测试大数据集（超过40个单词，会自动分页）
	err := GenExerciseSheet("test_resource", words, tempFile, false)
	if err != nil {
		t.Errorf("GenExerciseSheet() with large dataset 函数执行失败: %v", err)
	}
	
	// 检查文件是否创建成功
	if _, err := os.Stat(tempFile); os.IsNotExist(err) {
		t.Error("GenExerciseSheet() with large dataset 未生成输出文件")
	}
	
	// 清理临时文件
	os.Remove(tempFile)
}

func TestGenExerciseSheetWithShuffleCompare(t *testing.T) {
	// 创建一个单词列表用于测试打乱功能
	words := []string{
		"n.单词A",
		"v.单词B", 
		"adj.单词C",
		"adv.单词D",
		"pron.单词E",
	}

	// 生成两个临时文件，一个打乱，一个不打乱
	tempFileOrdered := "temp_ordered.xlsx"
	tempFileShuffled := "temp_shuffled.xlsx"
	
	// 生成不打乱的文件
	err := GenExerciseSheet("test_resource", words, tempFileOrdered, false)
	if err != nil {
		t.Errorf("GenExerciseSheet() ordered 函数执行失败: %v", err)
	}
	
	// 生成打乱的文件
	err = GenExerciseSheet("test_resource", words, tempFileShuffled, true)
	if err != nil {
		t.Errorf("GenExerciseSheet() shuffled 函数执行失败: %v", err)
	}
	
	// 检查两个文件是否都创建成功
	if _, err := os.Stat(tempFileOrdered); os.IsNotExist(err) {
		t.Error("GenExerciseSheet() ordered 未生成输出文件")
	}
	
	if _, err := os.Stat(tempFileShuffled); os.IsNotExist(err) {
		t.Error("GenExerciseSheet() shuffled 未生成输出文件")
	}
	
	// 清理临时文件
	os.Remove(tempFileOrdered)
	os.Remove(tempFileShuffled)
}

func TestGenExerciseSheetWithOptions(t *testing.T) {
	// 创建一个临时的单词列表
	words := []string{
		"n.测试单词1",
		"v.测试单词2",
		"adj.测试单词3",
	}

	// 生成临时文件
	tempFile := "temp_test_options.xlsx"
	
	// 测试带选项的生成函数
	opts := GenerateOptions{
		ShowPos:   true,
		WordCount: -1,
		Shuffle:   false,
	}
	generator := NewExerciseGenerator("test_resource", opts, words)
	err := generator.Generate(tempFile)
	if err != nil {
		t.Errorf("GenExerciseSheetWithOptions() 函数执行失败: %v", err)
	}
	
	// 检查文件是否创建成功
	if _, err := os.Stat(tempFile); os.IsNotExist(err) {
		t.Error("GenExerciseSheetWithOptions() 未生成输出文件")
	}
	
	// 清理临时文件
	os.Remove(tempFile)
}

func TestGenExerciseSheetSingleWord(t *testing.T) {
	// 创建一个单词的列表
	words := []string{"n.单个测试单词"}

	// 生成临时文件
	tempFile := "temp_test_single.xlsx"
	
	// 测试单个单词的情况
	err := GenExerciseSheet("test_resource", words, tempFile, false)
	if err != nil {
		t.Errorf("GenExerciseSheet() with single word 函数执行失败: %v", err)
	}
	
	// 检查文件是否创建成功
	if _, err := os.Stat(tempFile); os.IsNotExist(err) {
		t.Error("GenExerciseSheet() with single word 未生成输出文件")
	}
	
	// 清理临时文件
	os.Remove(tempFile)
}

func TestEngColStyles(t *testing.T) {
	// 创建Excel文件实例
	f := excelize.NewFile()
	sheet := "Sheet1"
	f.NewSheet(sheet)
	
	// 测试EngColStyTop函数
	generator := NewExcelGenerator()
	topStyle, err := generator.EngColStyTop(f)
	if err != nil {
		t.Errorf("EngColStyTop() 函数执行失败: %v", err)
	}
	if topStyle <= 0 {
		t.Error("EngColStyTop() 应该返回有效的样式ID")
	}
	
	// 测试EngColStyMid函数
	midStyle, err := generator.EngColStyMid(f)
	if err != nil {
		t.Errorf("EngColStyMid() 函数执行失败: %v", err)
	}
	if midStyle <= 0 {
		t.Error("EngColStyMid() 应该返回有效的样式ID")
	}
	
	// 测试EngColStyBot函数
	botStyle, err := generator.EngColStyBot(f)
	if err != nil {
		t.Errorf("EngColStyBot() 函数执行失败: %v", err)
	}
	if botStyle <= 0 {
		t.Error("EngColStyBot() 应该返回有效的样式ID")
	}
	
	// 测试样式是否正确应用了WrapText: false
	// 由于无法直接检查样式对象，我们测试函数不会报错
	f.Close()
}

func TestEngColStylesWithExerciseSheet(t *testing.T) {
	// 创建一个临时的单词列表
	words := []string{
		"n.测试单词1",
		"v.测试单词2",
	}

	// 生成临时文件
	tempFile := "temp_test_eng_styles.xlsx"
	
	// 测试正常情况
	err := GenExerciseSheet("test_resource", words, tempFile, false)
	if err != nil {
		t.Errorf("GenExerciseSheet() with EngColStyles 函数执行失败: %v", err)
	}
	
	// 检查文件是否创建成功
	if _, err := os.Stat(tempFile); os.IsNotExist(err) {
		t.Error("GenExerciseSheet() with EngColStyles 未生成输出文件")
	}
	
	// 清理临时文件
	os.Remove(tempFile)
}