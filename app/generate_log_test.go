package main

import (
	"fmt"
	"testing"
)

type SimMock struct {
	data map[string]map[string]string // In-memory representation of Excel data
}

func (s *SimMock) GetCellValue(sheet, cell string) (string, error) {
	if sheetData, ok := s.data[sheet]; ok {
		if cellValue, ok := sheetData[cell]; ok {
			return cellValue, nil
		}
	}
	return "", fmt.Errorf("Cell %s!%s is missing", sheet, cell)
}

func (s *SimMock) Close() error {
	return nil
}

func TestSimHour(t *testing.T) {
}
