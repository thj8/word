package main

import (
	"flag"
	"fmt"
	"os"
	"path/filepath"
	"strings"
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

	// 如果指定了单词数量，截取相应数量的单词
	if *wordCount > 0 && *wordCount < len(words) {
		if *shuffle {
			// 如果需要打乱顺序，在截取前先打乱整个列表
			words = utils.ShuffleWords(words)
		}
		// 截取指定数量的单词
		if *wordCount < len(words) {
			words = words[:*wordCount]
		}
	} else if *shuffle {
		// 如果不需要截断但需要打乱
		words = utils.ShuffleWords(words)
	}

	// 创建excel目录
	excelDir := "excel"
	if _, err := os.Stat(excelDir); os.IsNotExist(err) {
		err := os.Mkdir(excelDir, 0755)
		if err != nil {
			fmt.Printf("创建目录失败: %v\n", err)
			return
		}
	}

	// 定义数据集
	allWordsList := words

	// 清理资源名称，去除不适合作为文件名的字符
	cleanResourceName := utils.CleanFileName(resourceNameValue)

	// 添加参数信息到文件名
	filenameParts := []string{cleanResourceName}
	if !*showPos {
		filenameParts = append(filenameParts, "no_pos")
	}
	if *wordCount > 0 {
		filenameParts = append(filenameParts, fmt.Sprintf("%dwords", *wordCount))
	}
	if *shuffle {
		filenameParts = append(filenameParts, "shuffle")
	}

	filename := filepath.Join(excelDir, fmt.Sprintf("%s.xlsx", strings.Join(filenameParts, "_")))

	opts := tool.GenerateOptions{
		ShowPos:   *showPos,
		WordCount: *wordCount,
		Shuffle:   *shuffle,
	}
	if err := tool.GenerateExerciseSheet(allWordsList, filename, opts); err != nil {
		fmt.Printf("生成Excel文件时发生错误: %v\n", err)
	} else {
		fmt.Printf("成功生成Excel文件: %s\n", filename)
	}
}