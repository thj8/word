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
			name: "混合对象和字符串",
			rawWords: []interface{}{
				map[string]interface{}{"pos": "n.", "text": "cat"},
				"dog",
				map[string]interface{}{"pos": "v.", "text": "run"},
			},
			expected: []string{"n.cat", "dog", "v.run"},
		},
		{
			name: "只有text的对象",
			rawWords: []interface{}{
				map[string]interface{}{"text": "apple"},
			},
			expected: []string{"apple"},
		},
		{
			name: "只有pos的对象",
			rawWords: []interface{}{
				map[string]interface{}{"pos": "n."},
			},
			expected: []string{"n."},
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