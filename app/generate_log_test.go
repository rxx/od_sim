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

func newMockGameLog(sim *SimMock, actions ...ActionFunc) *GameLogCmd {
	return &GameLogCmd{
		currentHour: 0,
    sim: sim,
    actions: actions
	}
}

package main

import (
    "fmt"
    "strings"
    "testing"
    "time"
)

// ... (ExcelizeInterface, GameLogCmd, SimMock definitions)

func TestTickAction(t *testing.T) {
    testCases := []struct {
        name         string
        simData      map[string]map[string]string
        expected     string
        expectedErr  error
        currentHour  int
        simHour      int
    }{
        {
            name: "Valid Times and Date",
            simData: map[string]map[string]string{
                Overview: {"B15": "5/18/2024"},
                Imps:     {"BY1": "18:00", "BZ1": "00:00"},
            },
            expected:     "====== Protection Hour: 1 ( Local Time: 6:00:00 PM 5/18/2024 ) ( Domtime: 12:00:00 AM 5/19/2024 ) ======",
            expectedErr:  nil,
            currentHour:  0,
            simHour:      1,
        },
        {
            name: "Invalid Local Time Format",
            simData: map[string]map[string]string{
                Overview: {"B15": "5/18/2024"},
                Imps:     {"BY1": "invalid", "BZ1": "00:00"},
            },
            expected:     "",
            expectedErr:  fmt.Errorf("error parsing local time"),
            currentHour:  0,
            simHour:      1,
        },
        {
            name: "Invalid Date Format",
            simData: map[string]map[string]string{
                Overview: {"B15": "invalid"},
                Imps:     {"BY1": "18:00", "BZ1": "00:00"},
            },
            expected:     "",
            expectedErr:  fmt.Errorf("error parsing date"),
            currentHour:  0,
            simHour:      1,
        },
        // Add more test cases here for different scenarios
        // ... (e.g., missing time values, time overflow, incorrect hour order)
    }

    for _, tc := range testCases {
        t.Run(tc.name, func(t *testing.T) {
            mockSim := &SimMock{Data: tc.simData}
            glc := &GameLogCmd{
                currentHour: tc.currentHour,
                simHour:     tc.simHour,
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
        name         string
        simData      map[string]map[string]string
        expected     string
        expectedErr  error
        currentHour  int
    }{
        {
            name: "Draftrate Changed",
            simData: map[string]map[string]string{
                Military: {
                    "Y1":  "90%",
                    "Z0": "80%",
                },
            },
            expected:     "Draftrate changed to 90%",
            expectedErr:  nil,
            currentHour:  0,
        },
        {
            name: "Draftrate Changed (Decimal Format)",
            simData: map[string]map[string]string{
                Military: {
                    "Y2":  "0.85",
                    "Z1": "0.90",
                },
            },
            expected:     "Draftrate changed to 0.85",
            expectedErr:  nil,
            currentHour:  1,
        },
        {
            name: "Draftrate Unchanged",
            simData: map[string]map[string]string{
                Military: {
                    "Y3":  "0.75",
                    "Z2": "0.75",
                },
            },
            expected:     "",
            expectedErr:  nil,
            currentHour:  2,
        },
        {
            name: "Current Draftrate Empty",
            simData: map[string]map[string]string{
                Military: {
                    "Y4":  "",
                    "Z3": "0.60",
                },
            },
            expected:     "",
            expectedErr:  nil,
            currentHour:  3,
        },
        {
            name: "Error Reading Current Draftrate",
            simData: map[string]map[string]string{
                Military: {}, // Simulate missing cell
            },
            expected:     "",
            expectedErr:  fmt.Errorf("error reading current draftrate"),
            currentHour:  4,
        },
        {
            name: "Error Reading Previous Draftrate",
            simData: map[string]map[string]string{
                Military: {"Y5": "0.50"}, // Simulate missing previous cell
            },
            expected:     "",
            expectedErr:  fmt.Errorf("error reading previous draftrate"),
            currentHour:  4,
        },
    }

    for _, tc := range testCases {
        t.Run(tc.name, func(t *testing.T) {
            mockSim := &SimMock{Data: tc.simData}
            glc := &GameLogCmd{
                currentHour: tc.currentHour,
                simHour:     tc.currentHour + 1,
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
    testCases := []struct {
        name         string
        simData      map[string]map[string]string
        expected     string
        expectedErr  error
        currentHour  int
    }{
        {
            name: "Units and Draftees Released",
            simData: map[string]map[string]string{
                Military: {
                    "AX2": "Swordsmen",
                    "AX1": "10",
                    "AY2": "Archers",
                    "AY1": "5",
                    "AW1": "20",  
                },
            },
            expected:     "You successfully released 10 Swordsmen, 5 Archers\nYou successfully released 20 draftees into the peasantry\n",
            expectedErr:  nil,
            currentHour:  0,
        },
        {
            name: "Only Units Released",
            simData: map[string]map[string]string{
                Military: {
                    "AZ2": "Cavalry",
                    "AZ1": "3",
                },
            },
            expected:     "You successfully released 3 Cavalry\n",
            expectedErr:  nil,
            currentHour:  0,
        },
        {
            name: "Only Draftees Released",
            simData: map[string]map[string]string{
                Military: {
                    "AW1": "15",
                },
            },
            expected:     "You successfully released 15 draftees into the peasantry\n",
            expectedErr:  nil,
            currentHour:  0,
        },
        {
            name: "No Units or Draftees Released",
            simData: map[string]map[string]string{
                Military: {
                    "AX1": "0",
                    "AY1": "0",
                    "AW1": "0", 
                },
            },
            expected:     "",
            expectedErr:  nil,
            currentHour:  0,
        },
        {
            name: "Error Reading Unit Count",
            simData: map[string]map[string]string{
                Military: {
                    "AX2": "Swordsmen",
                    "AX1": "invalid", // Invalid unit count
                },
            },
            expected:     "",
            expectedErr:  nil, // Since no error handling is implemented in releaseUnitsAction
            currentHour:  0,
        },
        // Add more test cases here as needed...
    }

    for _, tc := range testCases {
        t.Run(tc.name, func(t *testing.T) {
            // Same test logic as in TestTickAction
            mockSim := &SimMock{Data: tc.simData}
            glc := &GameLogCmd{
                currentHour: tc.currentHour,
                simHour:     tc.currentHour + 1,
                sim:         mockSim,
            }

            result, err := glc.releaseUnitsAction()
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







