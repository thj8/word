package utils

import (
	"encoding/json"
	"io/ioutil"
	"math/rand"
	"path/filepath"
	"strings"
	"time"
)

// WordEntry 表示一个单词条目，包含词性、中文和英文
type WordEntry struct {
	Pos  string `json:"pos"`
	Text string `json:"text"`  // 中文翻译
	Word string `json:"word"`  // 英文单词
}

// Resources 存储所有资源
type Resources map[string][]string

// LoadAllResources 从resources目录加载所有JSON资源文件
func LoadAllResources() Resources {
	resources := make(Resources)
	
	// 尝试多种可能的路径
	possiblePaths := []string{"resources", "../resources", "../../resources"}
	
	var resourceDir string
	for _, path := range possiblePaths {
		if _, err := ioutil.ReadDir(path); err == nil {
			resourceDir = path
			break
		}
	}
	
	if resourceDir == "" {
		return resources
	}

	// 读取resources目录中的所有JSON文件
	files, err := ioutil.ReadDir(resourceDir)
	if err != nil {
		return resources
	}

	for _, file := range files {
		if strings.HasSuffix(file.Name(), ".json") {
			filePath := filepath.Join(resourceDir, file.Name())
			
			// 使用文件名（不含扩展名）作为资源名称
			resourceName := strings.TrimSuffix(file.Name(), ".json")
			
			data, err := ioutil.ReadFile(filePath)
			if err != nil {
				continue
			}

			var rawWords []interface{}
			if err := json.Unmarshal(data, &rawWords); err != nil {
				continue
			}
			
			// 将原始数据转换为字符串数组
			words := ConvertRawWordsToStrings(rawWords)
			
			// 将读取到的单词数组存储到资源映射中
			resources[resourceName] = words
		}
	}

	return resources
}

// ShuffleWords 随机打乱单词列表的顺序
func ShuffleWords(words []string) []string {
	r := rand.New(rand.NewSource(time.Now().UnixNano()))
	shuffled := make([]string, len(words))
	copy(shuffled, words)
	
	// 执行洗牌算法
	for i := len(shuffled) - 1; i > 0; i-- {
		j := r.Intn(i + 1)
		shuffled[i], shuffled[j] = shuffled[j], shuffled[i]
	}
	
	return shuffled
}

// ConvertRawWordsToStrings 将原始JSON数据转换为字符串数组
func ConvertRawWordsToStrings(rawWords []interface{}) []string {
	var words []string
	
	for _, item := range rawWords {
		switch v := item.(type) {
		case string:
			// 如果是字符串，直接添加
			words = append(words, v)
		case map[string]interface{}:
			// 如果是对象，组合pos和text字段
			pos, posOk := v["pos"].(string)
			text, textOk := v["text"].(string)
			word, wordOk := v["word"].(string)
			
			if posOk && textOk {
				words = append(words, pos+text) // 使用pos+text的组合
			} else if textOk {
				// 如果只有text字段，直接使用text
				words = append(words, text)
			} else if wordOk {
				// 如果只有word字段，使用word
				words = append(words, word)
			} else if posOk {
				// 如果只有pos字段，直接使用pos
				words = append(words, pos)
			}
		}
	}
	
	return words
}

// ExtractTextWithoutPos 从单词字符串中提取不包含pos的部分
func ExtractTextWithoutPos(word string) string {
	// 查找第一个点的位置，通常是pos和text之间的分隔符
	for i, char := range word {
		if char == '.' {
			// 返回点之后的部分（即text部分）
			if i+1 < len(word) {
				return word[i+1:]
			}
			break
		}
	}
	// 如果没有找到点，则返回原字符串
	return word
}