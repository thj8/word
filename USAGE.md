# 单词默写练习Excel生成器 - Go库使用指南

这是一个用于生成单词默写练习表格的Go库，可以将词汇数据导出为格式化的Excel文件，方便教师和学生使用。

## 安装

```bash
go get github.com/thj8/word
```

## 使用方法

### 基本使用

```go
package main

import (
	"fmt"
	"log"

	"github.com/thj8/word/lib"
)

func main() {
	// 显示可用资源
	availableResources := lib.GetAvailableResources()
	fmt.Println("Available resources:")
	for _, resource := range availableResources {
		fmt.Printf("  - %s\n", resource)
	}

	// 生成练习表
	err := lib.GenerateExerciseSheet("新概念青少版B", true, 10, false)
	if err != nil {
		log.Fatal(err)
	}

	fmt.Println("Excel file generated successfully!")
}
```

### 详细使用说明

#### 生成练习表

```go
// 生成练习表
// 参数: 资源名称, 是否显示词性, 单词数量(-1表示全部), 是否随机排序
err := lib.GenerateExerciseSheet("新概念青少版B", true, 10, false)
if err != nil {
    log.Fatal(err)
}
```

#### 获取可用资源列表

```go
// 获取所有可用的词汇资源
resources := lib.GetAvailableResources()
for _, resource := range resources {
    fmt.Println(resource)
}
```

！