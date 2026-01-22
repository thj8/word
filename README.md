# 单词默写练习Excel生成器

这是一个用于生成单词默写练习表格的Go库，可以将词汇数据导出为格式化的Excel文件，方便教师和学生使用。

## 安装

```bash
go get github.com/thj8/word
```

## 作为库使用

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

### API参考

#### 函数

- `lib.GenerateExerciseSheet(resourceName string, showPos bool, wordCount int, shuffle bool) error`
  - 根据指定的资源生成Excel练习表
  - `resourceName`: 资源名称（如"新概念青少版B"）
  - `showPos`: 是否显示词性
  - `wordCount`: 输出单词数量（-1表示全部）
  - `shuffle`: 是否随机排序
  - 返回错误信息

- `lib.GetAvailableResources() []string`
  - 获取所有可用的资源名称列表
  - 返回字符串切片

## 作为命令行工具使用

如果想直接使用命令行工具：

```bash
go run main.go -resource "新概念青少版B" -show-pos=true -count=10 -shuffle=false
```

## 文件结构

```
word/
├── main.go           # 命令行工具入口
├── lib/
│   └── lib.go        # 可导入的库
├── tool/
│   ├── generator.go  # Excel生成器核心实现
│   ├── generator_test.go # 单元测试
│   └── excel_generator.go  # Excel页面设置和格式化
├── utils/
│   ├── utils.go      # 工具函数，处理资源加载和数据转换
│   ├── utils_test.go # 单元测试
│   ├── file.go       # 文件操作相关工具函数
│   └── file_test.go  # 文件操作単元测试
├── resources/        # 词汇资源文件
│   ├── 新概念青少版B.json
│   ├── 译林三上.json
│   └── 新概念青少版A.json
├── go.mod
├── go.sum
└── README.md
```

## lib 文件夹说明

lib 文件夹包含库的公共接口，使得这个项目可以作为库被其他项目导入使用：

- **公共接口封装**：提供简单易用的 API，如 [GenerateExerciseSheet](file:///Users/sugar/Desktop/stu/word/lib/lib.go#L11-L43) 和 [GetAvailableResources](file:///Users/sugar/Desktop/stu/word/lib/lib.go#L46-L53)
- **内部实现隐藏**：隐藏复杂的内部逻辑，只暴露必要的功能给使用者
- **库与工具分离**：允许同一套代码既可以作为命令行工具使用，也可以作为库导入到其他项目中

## 贡献更多资源

我们欢迎社区贡献更多的词汇资源！如果您有以下类型的词汇资源，欢迎提交 PR：

- **英语教材词汇**：如剑桥英语、牛津英语、人教版英语等
- **其他语言词汇**：如法语、德语、日语、韩语等
- **专业词汇**：如医学英语、计算机英语、商务英语等
- **分级词汇**：如小学、初中、高中、大学四级、六级等

### 如何贡献

1. Fork 本仓库
2. 在 `resources/` 目录下添加您的 JSON 文件
3. 按照以下格式准备您的词汇数据：
```json
[
  {"pos": "n.", "text": "苹果"},
  {"pos": "v.", "text": "跑"},
  {"pos": "adj.", "text": "美丽的"}
]
```

4. 提交 Pull Request

我们非常感谢您的贡献！

## 资源文件格式

资源文件必须是JSON格式，包含单词数组：
```json
[
  {"pos": "n.", "text": "苹果"},
  {"pos": "v.", "text": "跑"},
  {"pos": "adj.", "text": "美丽的"}
]
```

## 功能特点

- 根据词汇列表自动生成格式化的Excel默写练习表
- 支持批量生成多个工作表的练习表格（每表最多40个单词）
- 自动设置页面布局和単元格样式
- 统一的文件命名规范
- 自动创建excel目录存放生成的文件
- 支持多资源管理
- 支持多种命令行选项，包括是否显示词性、输出单词数量和随机排序

## Excel表格格式

生成的Excel文件包含以下元素：

- 标题区域（"单词默写练习"和姓名/时间填写区）
- 表头（序号、中文、英文）
- 每个单词占用3行空间，便于书写
- 特殊边框样式（虚线分隔）
- A4纸张尺寸，适合打印

## 扩展性

程序采用模块化设计，可以轻松扩展：

- 在resources目录中添加更多JSON资源文件
- 修改样式设置
- 调整页面布局
- 更改文件命名规则

## 技术栈

- 语言: Go (Golang)
- Excel处理: excelize/v2
- 数据格式: JSON
- 架构: 模块化设计，便于维护和扩展

## 许可证

此项目仅供教育和个人使用。