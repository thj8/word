# 新概念青少版B单词默写练习Excel生成器

这是一个用于生成新概念青少版B系列单词默写练习表格的Go程序，可以将词汇数据导出为格式化的Excel文件，方便教师和学生使用。

## 功能特点

- 根据词汇列表自动生成格式化的Excel默写练习表
- 支持批量生成多个工作表的练习表格（每表最多40个单词）
- 自动设置页面布局和单元格样式
- 统一的文件命名规范
- 自动创建excel目录存放生成的文件

## 文件结构

```
word/
├── moxie.go          # 主程序入口，包含数据和生成逻辑
├── tool/
│   ├── generator.go  # Excel生成器核心实现
│   └── excel_generator.go  # Excel页面设置和格式化
└── README.md         # 项目说明文档
```

## 安装依赖

本项目使用`github.com/xuri/excelize/v2`库来操作Excel文件，可以通过以下命令安装：

```bash
go mod init word
go mod tidy 
```

## 使用方法

1. 在项目根目录运行程序：

```bash
go run moxie.go
```

2. 程序会根据内置的单词数据自动生成Excel文件，文件会存放在`excel`目录中：
   - 新概念-青少版B.xlsx（包含多个工作表，每表最多40个单词）

## 配置数据

在[moxie.go](./moxie.go)文件中可以找到所有单词数据，已合并为一个[allWords](file:///Users/sugar/Desktop/stu/word/moxie.go#L33-L97)列表：
- 程序会自动根据单词总数和每页最大40个单词的限制，计算并生成相应数量的工作表
- 每个工作表会自动命名为"Page1", "Page2", 等等

## Excel表格格式

生成的Excel文件包含以下元素：

- 标题区域（"单词默写练习"和姓名/时间填写区）
- 表头（序号、中文、英文）
- 每个单词占用3行空间，便于书写
- 特殊边框样式（虚线分隔）
- A4纸张尺寸，适合打印

## API说明

工具包提供了两个函数：

- `GenExerciseSheet(allWords []string, filename string, shuffle bool)` - 简化版函数，自动按每页40个单词分割数据并生成工作表，支持是否打乱顺序
- `GenExerSheetWithNames(datasets [][]string, sheetNames []string, filename string)` - 完整版函数，支持自定义数据分割和工作表名称

## 扩展性

程序采用模块化设计，可以轻松扩展：

- 添加更多单词到[allWords](file:///Users/sugar/Desktop/stu/word/moxie.go#L33-L97)列表
- 修改样式设置
- 调整页面布局
- 更改文件命名规则

## 技术栈

- 语言: Go (Golang)
- Excel处理: excelize/v2
- 架构: 模块化设计，便于维护和扩展

## 注意事项

- 程序会自动根据单词数量计算需要的工作表数量，每表最多40个单词
- 确保安装了正确的依赖库
- 生成的Excel文件会存放在`excel`目录中，会覆盖同名文件

## 许可证

此项目仅供教育和个人使用。