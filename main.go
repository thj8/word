package main

import (
	"fmt"
	"os"
	"path/filepath"

	"word/tool"
	"word/utils"
)

func main() {
	// 检查命令行参数
	if len(os.Args) < 2 {
		fmt.Println("用法: go run main.go <资源名称>")
		fmt.Println("可用资源:")
		resources := utils.LoadAllResources()
		for resourceName := range resources {
			fmt.Printf("  - %s\n", resourceName)
		}
		return
	}

	resourceName := os.Args[1]
	
	// 加载资源
	resources := utils.LoadAllResources()
	
	// 检查请求的资源是否存在
	words, exists := resources[resourceName]
	if !exists {
		fmt.Printf("错误: 找不到资源 '%s'\n", resourceName)
		fmt.Println("可用资源:")
		for resourceName := range resources {
			fmt.Printf("  - %s\n", resourceName)
		}
		return
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
	cleanResourceName := utils.CleanFileName(resourceName)

	filename := filepath.Join(excelDir, fmt.Sprintf("%s.xlsx", cleanResourceName))

	if err := tool.GenExerciseSheet(allWordsList, filename, false); err != nil {
		fmt.Printf("生成Excel文件时发生错误: %v\n", err)
	} else {
		fmt.Printf("成功生成Excel文件: %s\n", filename)
	}
}