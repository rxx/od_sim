package main

import (
	"fmt"
	"strings"
	"testing"

	"github.com/xuri/excelize/v2"
)

type SimMock struct {
	Data map[string]map[string]string // In-memory representation of Excel data
}

func (s *SimMock) GetCellValue(sheet, cell string, _ ...excelize.Options) (string, error) {
	if sheetData, ok := s.Data[sheet]; ok {
		if cellValue, ok := sheetData[cell]; ok {
			return cellValue, nil
		}
	}
	return "", fmt.Errorf("Cell %s!%s is missing", sheet, cell)
}

func (s *SimMock) Close() error {
	return nil
}

func newMockGameLog(sim *SimMock, actions ...ActionFunc) *GameLogCmd {
	return &GameLogCmd{
		currentHour: 0,
		sim:         sim,
		actions:     actions,
	}
}

// deepCopyAndMergeMaps creates a deep copy of the source map
// and merges the data from the override map into it.
func deepCopyAndMergeMaps(src map[string]map[string]string, override map[string]map[string]string) map[string]map[string]string {
	result := make(map[string]map[string]string)
	for sheet, cellValues := range src {
		result[sheet] = make(map[string]string)
		for cell, value := range cellValues {
			result[sheet][cell] = value
		}
	}
	for sheet, cellValues := range override {
		if _, ok := result[sheet]; !ok {
			result[sheet] = make(map[string]string)
		}
		for cell, value := range cellValues {
			result[sheet][cell] = value
		}
	}
	return result
}

func TestTickAction(t *testing.T) {
	testCases := []struct {
		name        string
		simData     map[string]map[string]string
		expected    string
		expectedErr error
		currentHour int
		simHour     int
	}{
		{
			name: "Valid Times and Date",
			simData: map[string]map[string]string{
				Overview: {"B15": "5/18/2024"},
				Imps:     {"BY4": "18:00", "BZ4": "00:00"},
			},
			expected:    "====== Protection Hour: 1 ( Local Time: 6:00:00 PM 5/18/2024 ) ( Domtime: 12:00:00 AM 5/18/2024 ) ======",
			expectedErr: nil,
		},
		{
			name: "Invalid Dom Time Format",
			simData: map[string]map[string]string{
				Overview: {"B15": "5/18/2024"},
				Imps:     {"BY4": "18:00", "BZ4": "invalid"},
			},
			expected:    "",
			expectedErr: fmt.Errorf("error parsing dom time"),
		},

		{
			name: "Invalid Local Time Format",
			simData: map[string]map[string]string{
				Overview: {"B15": "5/18/2024"},
				Imps:     {"BY4": "invalid", "BZ4": "00:00"},
			},
			expected:    "",
			expectedErr: fmt.Errorf("error parsing local time"),
		},
		{
			name: "Invalid Date Format",
			simData: map[string]map[string]string{
				Overview: {"B15": "invalid"},
				Imps:     {"BY4": "18:00", "BZ4": "00:00"},
			},
			expected:    "",
			expectedErr: fmt.Errorf("error parsing date"),
		},
	}

	for _, tc := range testCases {
		t.Run(tc.name, func(t *testing.T) {
			mockSim := &SimMock{Data: tc.simData}
			glc := &GameLogCmd{
				currentHour: 0,
				simHour:     4,
				sim:         mockSim,
			}

			result, err := glc.tickAction()

			if err != nil && tc.expectedErr == nil {
				t.Errorf("Unexpected error: %v", err)
			} else if err == nil && tc.expectedErr != nil {
				t.Errorf("Expected error, but got none")
			} else if err != nil && tc.expectedErr != nil && !strings.Contains(err.Error(), tc.expectedErr.Error()) {
				t.Errorf("Incorrect error message: got %q, want %q", err, tc.expectedErr)
			}
			if result != tc.expected {
				t.Errorf("Incorrect result: got %q, want %q", result, tc.expected)
			}
		})
	}
}

func TestDraftRateAction(t *testing.T) {
	testCases := []struct {
		name        string
		simData     map[string]map[string]string
		expected    string
		expectedErr error
		currentHour int
	}{
		{
			name: "Draftrate Changed from 80% to 90%",
			simData: map[string]map[string]string{
				Military: {
					"Y5": "90%",
					"Z4": "80%",
				},
			},
			expected:    "Draftrate changed to 90%",
			expectedErr: nil,
			currentHour: 1,
		},
		{
			name: "Draftrate Changed from blank",
			simData: map[string]map[string]string{
				Military: {
					"Y5": "90%",
					"Z4": "",
				},
			},
			expected:    "Draftrate changed to 90%",
			expectedErr: nil,
			currentHour: 1,
		},
		{
			name: "Draftrate Unchanged",
			simData: map[string]map[string]string{
				Military: {
					"Y5": "90%",
					"Z4": "90%",
				},
			},
			expected:    "",
			expectedErr: nil,
			currentHour: 1,
		},
		{
			name: "Current Draftrate Empty",
			simData: map[string]map[string]string{
				Military: {
					"Y5": "",
					"Z4": "90%",
				},
			},
			expected:    "",
			expectedErr: nil,
			currentHour: 1,
		},
		{
			name: "Error Reading Current Draftrate",
			simData: map[string]map[string]string{
				Military: {}, // Simulate missing cell
			},
			expected:    "",
			expectedErr: fmt.Errorf("error reading current draftrate"),
			currentHour: 1,
		},
		{
			name: "Error Reading Previous Draftrate",
			simData: map[string]map[string]string{
				Military: {"Y5": "90%"}, // Simulate missing previous cell
			},
			expected:    "",
			expectedErr: fmt.Errorf("error reading previous draftrate"),
			currentHour: 1,
		},
	}

	for _, tc := range testCases {
		t.Run(tc.name, func(t *testing.T) {
			mockSim := &SimMock{Data: tc.simData}
			glc := &GameLogCmd{
				currentHour: tc.currentHour,
				simHour:     tc.currentHour + 4,
				sim:         mockSim,
			}

			result, err := glc.draftRateAction()
			if err != nil && tc.expectedErr == nil {
				t.Errorf("Unexpected error: %v", err)
			} else if err == nil && tc.expectedErr != nil {
				t.Errorf("Expected error, but got none")
			} else if err != nil && tc.expectedErr != nil && !strings.Contains(err.Error(), tc.expectedErr.Error()) {
				t.Errorf("Incorrect error message: got %q, want %q", err, tc.expectedErr)
			}
			if result != tc.expected {
				t.Errorf("Incorrect result: got %q, want %q", result, tc.expected)
			}
		})
	}
}

func TestReleaseUnitsAction(t *testing.T) {
	releaseMilitaryMap := map[string]map[string]string{
		Military: {
			"AX2": "Spearman",
			"AY2": "Archer",
			"AZ2": "Knight",
			"BA2": "Cavalry",
			"BB2": "Spies",
			"BC2": "Archspies",
			"BD2": "Wizards",
			"BE2": "Archmages",
			"AX4": "",
			"AY4": "",
			"AZ4": "",
			"BA4": "",
			"BB4": "",
			"BC4": "",
			"BD4": "",
			"BE4": "",
			"AW4": "",
		},
	}

	testCases := []struct {
		name        string
		simData     map[string]map[string]string
		expected    string
		expectedErr error
		currentHour int
	}{
		{
			name: "Units and Draftees Released",
			simData: map[string]map[string]string{
				Military: {
					"AX4": "10",
					"AY4": "5",
					"AW4": "20",
				},
			},
			expected:    "You successfully released 10 Spearman, 5 Archer\nYou successfully released 20 draftees into the peasantry\n",
			expectedErr: nil,
			currentHour: 0,
		},
		{
			name: "Spies and wizards release",
			simData: map[string]map[string]string{
				Military: {
					"BB4": "10",
					"BC4": "5",
					"BD4": "20",
					"BE4": "10",
				},
			},
			expected:    "You successfully released 10 Spies, 5 Archspies, 20 Wizards, 10 Archmages\n",
			expectedErr: nil,
			currentHour: 0,
		},

		{
			name: "One Unit Released",
			simData: map[string]map[string]string{
				Military: {
					"AZ4": "3",
				},
			},
			expected:    "You successfully released 3 Knight\n",
			expectedErr: nil,
			currentHour: 0,
		},
		{
			name: "Only Draftees Released",
			simData: map[string]map[string]string{
				Military: {
					"AW4": "15",
				},
			},
			expected:    "You successfully released 15 draftees into the peasantry\n",
			expectedErr: nil,
			currentHour: 0,
		},
		{
			name: "Zero Units or Draftees Released",
			simData: map[string]map[string]string{
				Military: {
					"AX4": "0",
					"AY4": "0",
					"AW4": "0",
				},
			},
			expected:    "",
			expectedErr: nil,
			currentHour: 0,
		},
		{
			name: "Empty Units or Draftees Released",
			simData: map[string]map[string]string{
				Military: {
					"AX4": "",
					"AY4": "",
					"AW4": "",
				},
			},
			expected:    "",
			expectedErr: nil,
			currentHour: 0,
		},

		{
			name: "Error Reading Unit Count",
			simData: map[string]map[string]string{
				Military: {
					"AX4": "invalid", // Invalid unit count
				},
			},
			expected:    "",
			expectedErr: fmt.Errorf("error parsing unit value"),
			currentHour: 0,
		},
	}

	for _, tc := range testCases {
		t.Run(tc.name, func(t *testing.T) {
			mergedData := deepCopyAndMergeMaps(releaseMilitaryMap, tc.simData)
			mockSim := &SimMock{Data: mergedData}
			glc := &GameLogCmd{
				currentHour: tc.currentHour,
				simHour:     tc.currentHour + 4,
				sim:         mockSim,
			}

			result, err := glc.releaseUnitsAction()
			if err != nil && tc.expectedErr == nil {
				t.Errorf("Unexpected error: %v", err)
			} else if err == nil && tc.expectedErr != nil {
				t.Errorf("Expected error: %v, but got none", tc.expectedErr)
			} else if err != nil && tc.expectedErr != nil && !strings.Contains(err.Error(), tc.expectedErr.Error()) {
				t.Errorf("Incorrect error message: got %q, want %q", err, tc.expectedErr)
			}
			if result != tc.expected {
				t.Errorf("Incorrect result: got %q, want %q", result, tc.expected)
			}
		})
	}
}
