package lib

import (
	"fmt"
	"os"

	"github.com/thj8/word/tool"
	"github.com/thj8/word/utils"
)

// GenerateExerciseSheet generates an Excel file containing vocabulary exercises based on the specified resource name and options
func GenerateExerciseSheet(resourceName string, showPos bool, wordCount int, shuffle bool) error {
	// 加载资源
	resources := utils.LoadAllResources()

	// 检查请求的资源是否存在
	words, exists := resources[resourceName]
	if !exists {
		return fmt.Errorf("错误: 找不到资源 '%s'", resourceName)
	}

	// 创建excel目录
	excelDir := "excel"
	if _, err := os.Stat(excelDir); os.IsNotExist(err) {
		err := os.Mkdir(excelDir, 0o755)
		if err != nil {
			return fmt.Errorf("创建目录失败: %v", err)
		}
	}

	opts := tool.GenerateOptions{
		ShowPos:   showPos,
		WordCount: wordCount,
		Shuffle:   shuffle,
	}
	generator := tool.NewExerciseGenerator(resourceName, opts, words)
	filename := generator.GenerateFilename()
	if err := generator.GenerateAuto(); err != nil {
		return fmt.Errorf("生成Excel文件时发生错误: %v", err)
	} else {
		fmt.Printf("成功生成Excel文件: %s\n", filename)
		return nil
	}
}

// GetAvailableResources returns a list of available vocabulary resources
func GetAvailableResources() []string {
	resources := utils.LoadAllResources()
	names := make([]string, 0, len(resources))
	for name := range resources {
		names = append(names, name)
	}
	return names
}