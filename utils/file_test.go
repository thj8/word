package utils

import "testing"

func TestCleanFileName(t *testing.T) {
	tests := []struct {
		name     string
		input    string
		expected string
	}{
		{
			name:     "正常名称",
			input:    "新概念青少版B",
			expected: "新概念青少版B",
		},
		{
			name:     "包含斜杠",
			input:    "新概念/青少版B",
			expected: "新概念_青少版B",
		},
		{
			name:     "包含反斜杠",
			input:    "新概念\\青少版B",
			expected: "新概念_青少版B",
		},
		{
			name:     "包含冒号",
			input:    "新概念:青少版B",
			expected: "新概念_青少版B",
		},
		{
			name:     "包含星号",
			input:    "新概念*青少版B",
			expected: "新概念_青少版B",
		},
		{
			name:     "包含问号",
			input:    "新概念?青少版B",
			expected: "新概念_青少版B",
		},
		{
			name:     "包含引号",
			input:    "新概念\"青少版B",
			expected: "新概念_青少版B",
		},
		{
			name:     "包含小于号",
			input:    "新概念<青少版B",
			expected: "新概念_青少版B",
		},
		{
			name:     "包含大于号",
			input:    "新概念>青少版B",
			expected: "新概念_青少版B",
		},
		{
			name:     "包含竖线",
			input:    "新概念|青少版B",
			expected: "新概念_青少版B",
		},
		{
			name:     "包含所有非法字符",
			input:    "文件:名/路径\\测试*文件?名\"测试<文件>名|测试",
			expected: "文件_名_路径_测试_文件_名_测试_文件_名_测试",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := CleanFileName(tt.input)
			if result != tt.expected {
				t.Errorf("CleanFileName(%q) = %q, want %q", tt.input, result, tt.expected)
			}
		})
	}
}