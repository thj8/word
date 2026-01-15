package utils

import (
	"encoding/json"
	"io/ioutil"
	"path/filepath"
	"strings"
)

// WordEntry 表示一个单词条目，包含词性和文本
type WordEntry struct {
	Pos  string `json:"pos"`
	Text string `json:"text"`
}

// Resources 存储所有资源
type Resources map[string][]string

// LoadAllResources 从resources目录加载所有JSON资源文件
func LoadAllResources() Resources {
	resources := make(Resources)
	
	// 读取resources目录中的所有JSON文件
	files, err := ioutil.ReadDir("resources")
	if err != nil {
		return resources
	}

	for _, file := range files {
		if strings.HasSuffix(file.Name(), ".json") {
			filePath := filepath.Join("resources", file.Name())
			
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
			
			if posOk && textOk {
				words = append(words, pos+text)
			} else if textOk {
				// 如果只有text字段，直接使用text
				words = append(words, text)
			} else if posOk {
				// 如果只有pos字段，直接使用pos
				words = append(words, pos)
			}
		}
	}
	
	return words
}