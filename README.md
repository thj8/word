# 新概念青少版B单词默写练习Excel生成器

这是一个用于生成新概念青少版B系列单词默写练习表格的Go程序，可以将词汇数据导出为格式化的Excel文件，方便教师和学生使用。

## 功能特点

- 根据词汇列表自动生成格式化的Excel默写练习表
- 支持批量生成多个工作表的练习表格（每表最多40个单词）
- 自动设置页面布局和单元格样式
- 统一的文件命名规范
- 自动创建excel目录存放生成的文件
- 支持多资源管理，可通过命令行参数指定生成哪个资源
- 支持多种命令行选项，包括是否显示词性、输出单词数量和随机排序

## 文件结构

```
word/
├── main.go           # 主程序入口，处理命令行参数和基本流程
├── utils/
│   ├── utils.go      # 工具函数，处理资源加载和数据转换
│   ├── utils_test.go # 单元测试
│   ├── file.go       # 文件操作相关工具函数
│   └── file_test.go  # 文件操作单元测试
├── resources/        # 资源目录，存储所有词汇数据
│   └── 新概念青少版B.json  # 新概念青少版B词汇资源
├── tool/
│   ├── generator.go  # Excel生成器核心实现
│   ├── generator_test.go # 单元测试
│   └── excel_generator.go  # Excel页面设置和格式化
├── go.mod            # Go模块定义
├── go.sum            # Go依赖校验和
└── README.md         # 项目说明文档
```

## 安装依赖

本项目使用`github.com/xuri/excelize/v2`库来操作Excel文件，可以通过以下命令安装：

```bash
go mod tidy 
```

## 使用方法

1. 在项目根目录运行程序，使用命令行参数：

```bash
# 基本用法
go run main.go -resource "新概念青少版B"

# 显示所有选项的帮助
go run main.go -help

# 显示词性，输出20个单词，不打乱顺序
go run main.go -resource "新概念青少版B" -show-pos=true -count=20 -shuffle=false

# 不显示词性，输出10个单词，打乱顺序
go run main.go -resource "新概念青少版B" -show-pos=false -count=10 -shuffle=true
```

2. 程序会根据指定的资源名称生成Excel文件，文件会存放在`excel`目录中

## 命令行选项

- `-resource`: 指定资源名称（必填）
- `-show-pos`: 是否显示词性（默认true）
- `-count`: 输出单词个数（默认-1表示全部）
- `-shuffle`: 是否随机乱序（默认false）
- `-help`: 显示帮助信息

## 添加新资源

要添加新的资源，只需在`resources`目录中创建一个新的JSON文件，文件名即为资源名称，内容为单词数组：

例如，创建`初中英语词汇.json`：
```json
[
  {"pos": "n.", "text": "单词1"},
  {"pos": "v.", "text": "动词1"},
  ...
]
```

程序会自动加载`resources`目录中的所有JSON文件，文件名（不含扩展名）作为资源名称。

### JSON格式规范

资源文件必须遵循以下格式规范：

- 文件必须是有效的JSON格式
- 文件内容是一个数组，包含多个词汇对象
- 每个词汇对象必须包含以下字段：
  - `pos`: 词性（part of speech），如 "n."（名词）、"v."（动词）等
  - `text`: 词汇的具体含义或中文释义

### 词性（POS）属性详解

| 词性缩写 | 说明 | 示例 |
|---------|------|------|
| `n.` | 名词 (noun)：表示人、事物、地点或抽象概念的名称 | 学生、桌子、北京 |
| `v.` | 动词 (verb)：表示动作或状态 | 跑、学习、是 |
| `adj.` | 形容词 (adjective)：修饰名词或代词，描述人或事物的特征 | 美丽的、高的、聪明的 |
| `adv.` | 副词 (adverb)：修饰动词、形容词或其他副词 | 快速地、非常、很 |
| `pron.` | 代词 (pronoun)：代替名词或名词短语 | 我、你、他们 |
| `prep.` | 介词 (preposition)：表示名词或代词与其他词的关系 | 在、关于、为了 |
| `conj.` | 连词 (conjunction)：连接词、短语或句子 | 和、但是、或者 |
| `int.` | 感叹词 (interjection)：表达情感或感叹 | 哎呀、哦、哈哈 |
| `det.` | 限定词 (determiner)：限定名词的意义范围 | 这、那、一些 |
| `phr.` | 短语 (phrase)：固定搭配或常用语句 | 早上好、生日快乐 |
| `art.` | 冠词 (article)：用在名词前帮助说明名词的含义 | 一个、这个 |
| `num.` | 数词 (numeral)：表示数量或顺序 | 一、第一、许多 |

## Excel表格格式

生成的Excel文件包含以下元素：

- 标题区域（"单词默写练习"和姓名/时间填写区）
- 表头（序号、中文、英文）
- 每个单词占用3行空间，便于书写
- 特殊边框样式（虚线分隔）
- A4纸张尺寸，适合打印

## API说明

工具包提供了以下主要函数和结构体：

- `ExerciseGenerator` - 练习表生成器结构体，封装了所有生成逻辑
  - `NewExerciseGenerator(resourceName string, opts GenerateOptions, originalWords []string) *ExerciseGenerator` - 创建新的练习表生成器
  - `Generate(filename string) error` - 根据指定文件名生成练习表
  - `GenerateAuto() error` - 自动根据资源名称生成练习表
  - `GenerateFilename() string` - 根据资源名称和选项生成Excel文件名
- `GenerateOptions` - 定义生成Excel文件的选项结构体，包含以下字段：
  - `ShowPos bool` - 是否显示词性（默认 true）
  - `WordCount int` - 输出单词个数，-1 表示全部（默认 -1）
  - `Shuffle bool` - 是否随机乱序（默认 false）
- `GenExerciseSheet(resourceName string, allWords []string, filename string, shuffle bool) error` - 简化版函数（保持向后兼容），默认显示词性且输出全部单词，支持是否打乱顺序
- `GenerateExcelFilename(resourceName string, showPos bool, wordCount int, shuffle bool) string` - 根据资源名称和选项生成Excel文件名
- 样式相关函数：
  - `EngColStyTop(f *excelize.File) (int, error)` - 创建英文列顶部样式，禁用自动换行
  - `EngColStyMid(f *excelize.File) (int, error)` - 创建英文列中间样式，禁用自动换行
  - `EngColStyBot(f *excelize.File) (int, error)` - 创建英文列底部样式，禁用自动换行

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

## 注意事项

- 程序会自动根据单词数量计算需要的工作表数量，每表最多40个单词
- 确保安装了正确的依赖库
- 生成的Excel文件会存放在`excel`目录中，会覆盖同名文件
- 资源名称需要与JSON文件名（不含扩展名）完全匹配

## 许可证

此项目仅供教育和个人使用。