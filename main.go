package main

import (
	"flag"
	"fmt"
	"os"

	"word/tool"
	"word/utils"
)

func main() {
	// 定义命令行参数
	var (
		resourceName = flag.String("resource", "", "资源名称（必填）")
		showPos      = flag.Bool("show-pos", true, "是否显示词性（默认true）")
		wordCount    = flag.Int("count", -1, "输出单词个数（默认-1表示全部）")
		shuffle      = flag.Bool("shuffle", false, "是否随机乱序（默认false）")
		help         = flag.Bool("help", false, "显示帮助信息")
	)

	// 解析命令行参数
	flag.Parse()

	// 如果设置了-help参数或没有提供资源名称，显示帮助信息
	if *help || *resourceName == "" {
		fmt.Println("用法: go run main.go [选项]")
		fmt.Println("选项:")
		flag.PrintDefaults()
		fmt.Println("\n可用资源:")
		resources := utils.LoadAllResources()
		for resourceName := range resources {
			fmt.Printf("  - %s\n", resourceName)
		}
		return
	}

	resourceNameValue := *resourceName

	// 加载资源
	resources := utils.LoadAllResources()

	// 检查请求的资源是否存在
	words, exists := resources[resourceNameValue]
	if !exists {
		fmt.Printf("错误: 找不到资源 '%s'\n", resourceNameValue)
		fmt.Println("可用资源:")
		for resourceName := range resources {
			fmt.Printf("  - %s\n", resourceName)
		}
		return
	}

	// 创建excel目录
	excelDir := "excel"
	if _, err := os.Stat(excelDir); os.IsNotExist(err) {
		err := os.Mkdir(excelDir, 0o755)
		if err != nil {
			fmt.Printf("创建目录失败: %v\n", err)
			return
		}
	}

	opts := tool.GenerateOptions{
		ShowPos:   *showPos,
		WordCount: *wordCount,
		Shuffle:   *shuffle,
	}
	generator := tool.NewExerciseGenerator(resourceNameValue, opts, words)
	filename := generator.GenerateFilename()
	if err := generator.GenerateAuto(); err != nil {
		fmt.Printf("生成Excel文件时发生错误: %v\n", err)
	} else {
		fmt.Printf("成功生成Excel文件: %s\n", filename)
	}
}
