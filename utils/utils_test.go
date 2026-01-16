package utils

import (
	"reflect"
	"testing"
)

func TestConvertRawWordsToStrings(t *testing.T) {
	tests := []struct {
		name     string
		rawWords []interface{}
		expected []string
	}{
		{
			name:     "纯字符串数组",
			rawWords: []interface{}{"hello", "world"},
			expected: []string{"hello", "world"},
		},
		{
			name: "包含pos和text的对象",
			rawWords: []interface{}{
				map[string]interface{}{"pos": "n.", "text": "cat"},
				map[string]interface{}{"pos": "v.", "text": "run"},
			},
			expected: []string{"n.cat", "v.run"},
		},
		{
			name: "仅包含text的对象（pos为空或缺失）",
			rawWords: []interface{}{
				map[string]interface{}{"text": "apple"},
				map[string]interface{}{"text": "banana"},
			},
			expected: []string{"apple", "banana"},
		},
		{
			name: "混合包含pos+text和仅有text的对象",
			rawWords: []interface{}{
				map[string]interface{}{"pos": "n.", "text": "dog"},
				map[string]interface{}{"text": "elephant"},
				map[string]interface{}{"pos": "adj.", "text": "big"},
			},
			expected: []string{"n.dog", "elephant", "adj.big"},
		},
		{
			name: "仅有word字段的对象",
			rawWords: []interface{}{
				map[string]interface{}{"word": "hello"},
				map[string]interface{}{"word": "world"},
			},
			expected: []string{"hello", "world"},
		},
		{
			name: "仅有pos字段的对象",
			rawWords: []interface{}{
				map[string]interface{}{"pos": "n."},
				map[string]interface{}{"pos": "v."},
			},
			expected: []string{"n.", "v."},
		},
		{
			name: "包含pos、text和word的对象（应使用pos+text）",
			rawWords: []interface{}{
				map[string]interface{}{"pos": "n.", "text": "猫", "word": "cat"},
				map[string]interface{}{"pos": "v.", "text": "跑", "word": "run"},
			},
			expected: []string{"n.猫", "v.跑"}, // 应该使用pos+text，而不是word
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := ConvertRawWordsToStrings(tt.rawWords)
			if !reflect.DeepEqual(result, tt.expected) {
				t.Errorf("ConvertRawWordsToStrings() = %v, want %v", result, tt.expected)
			}
		})
	}
}

func TestShuffleWords(t *testing.T) {
	// 创建一个单词列表
	original := []string{"apple", "banana", "cherry", "date", "elderberry"}
	
	// 复制原始列表以进行比较
	originalCopy := make([]string, len(original))
	copy(originalCopy, original)
	
	// 打乱列表
	shuffled := ShuffleWords(original)
	
	// 验证打乱后的列表包含相同的元素
	if len(shuffled) != len(originalCopy) {
		t.Errorf("ShuffleWords() returned wrong length: got %d, want %d", len(shuffled), len(originalCopy))
	}
	
	// 检查所有原始元素是否仍在列表中
	for _, origItem := range originalCopy {
		found := false
		for _, shuffledItem := range shuffled {
			if origItem == shuffledItem {
				found = true
				break
			}
		}
		if !found {
			t.Errorf("ShuffleWords() lost an item: %s", origItem)
		}
	}
	
	// 由于random的性质，我们不能保证每次都会改变顺序，所以只做基本验证
}