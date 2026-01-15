package utils

import "strings"

// CleanFileName 清理文件名，去除不适合作为文件名的字符
func CleanFileName(name string) string {
	cleanName := strings.ReplaceAll(name, "/", "_")
	cleanName = strings.ReplaceAll(cleanName, "\\", "_")
	cleanName = strings.ReplaceAll(cleanName, ":", "_")
	cleanName = strings.ReplaceAll(cleanName, "*", "_")
	cleanName = strings.ReplaceAll(cleanName, "?", "_")
	cleanName = strings.ReplaceAll(cleanName, "\"", "_")
	cleanName = strings.ReplaceAll(cleanName, "<", "_")
	cleanName = strings.ReplaceAll(cleanName, ">", "_")
	cleanName = strings.ReplaceAll(cleanName, "|", "_")
	return cleanName
}