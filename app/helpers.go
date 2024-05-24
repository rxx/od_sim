package main

import (
	"fmt"
	"path/filepath"
	"runtime"
	"strings"
)

func debugLog(values ...interface{}) {
	if !debugEnabled {
		return
	}

	formattedValues := make([]interface{}, len(values))
	for i, value := range values {
		switch v := value.(type) {
		case string:
			if strings.TrimSpace(v) == "" {
				formattedValues[i] = "[empty]"
			} else {
				formattedValues[i] = v
			}
		case int, int8, int16, int32, int64, uint, uint8, uint16, uint32, uint64, float32, float64:
			if v == 0 { // Check for zero value of numeric types
				formattedValues[i] = "[empty]"
			} else {
				formattedValues[i] = v
			}
		default:
			formattedValues[i] = value // For other types, keep as is
		}
	}

	pc, file, line, _ := runtime.Caller(1)
	funcName := runtime.FuncForPC(pc).Name()

	fmt.Printf("--- DEBUG on [%s:%s:%d] ---\n", filepath.Base(file), funcName, line)
	fmt.Println(formattedValues...)
	fmt.Println("--- ^_^ ---")
}
