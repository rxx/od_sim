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
			expectedErr: fmt.Errorf("error reading unit value"),
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

_______

func TestParseMagicActions(t *testing.T) {
    testCases := []struct {
        name         string
        simData      map[string]map[string]string
        expected     string
        expectedErr  error
        simHour      int
    }{
        {
            name: "Multiple Spells Cast with Explore",
            simData: map[string]map[string]string{
                ExploreSheet: {
                    fmt.Sprintf("S%d", 1): "1", // Explore is active
                },
                MagicSheet: {
                    fmt.Sprintf("B%d", 1):  "50",  // Mana available
                    fmt.Sprintf("G%d", 1):  "1",   // Gaia's Watch cast
                    fmt.Sprintf("H%d", 1):  "1",   // Mining Strength cast
                    fmt.Sprintf("I%d", 1):  "1",   // Ares' Call cast
                },
            },
            expected: "Your wizards successfully cast Gaia's Watch at a cost of 80 mana.\n" +
                "Your wizards successfully cast Mining Strength at a cost of 80 mana.\n" +
                "Your wizards successfully cast Ares' Call at a cost of 200 mana.\n",
            expectedErr: nil,
            simHour:     1,
        },
        {
            name: "Single Spell Cast Without Explore",
            simData: map[string]map[string]string{
                MagicSheet: {
                    fmt.Sprintf("B%d", 2): "30",  // Mana available
                    fmt.Sprintf("J%d", 2): "1",   // Midas' Touch cast
                },
            },
            expected:     "Your wizards successfully cast Midas Touch at a cost of 75 mana.\n",
            expectedErr:  nil,
            simHour:     2,
        },
        {
            name: "No Spells Cast",
            simData: map[string]map[string]string{
                ExploreSheet: {
                    fmt.Sprintf("S%d", 3): "1", // Explore is active
                },
                MagicSheet: {
                    fmt.Sprintf("B%d", 3): "60", // Mana available
                },
            },
            expected:     "", // No message expected
            expectedErr:  nil,
            simHour:     3,
        },
        {
            name: "Error Reading Explore Cell",
            simData:      map[string]map[string]string{}, // Missing Explore data
            expected:     "",
            expectedErr:  fmt.Errorf("error reading cell Explore!S4: %w", &excelize.ErrCellNotFound{}), // Adjust error type if needed
            simHour:     4,
        },
        {
            name: "Error Reading Magic Cell",
            simData: map[string]map[string]string{
                ExploreSheet: {
                    fmt.Sprintf("S%d", 5): "1", // Explore is active
                },
            }, // Missing Magic data
            expected:     "",
            expectedErr:  fmt.Errorf("error reading cell Magic!G5: %w", &excelize.ErrCellNotFound{}), // Adjust error type
            simHour:     5,
        },
        {
            name: "Error Converting Mana to Integer",
            simData: map[string]map[string]string{
                ExploreSheet: {
                    fmt.Sprintf("S%d", 6): "1", // Explore is active
                },
                MagicSheet: {
                    fmt.Sprintf("B%d", 6):  "invalid", // Invalid mana value
                    fmt.Sprintf("G%d", 6):  "1",        // Gaia's Watch cast
                },
            },
            expected:     "",
            expectedErr:  fmt.Errorf("error converting cell Magic!B6 value to int"), 
            simHour:     6,
        },
    }

    for _, tc := range testCases {
        t.Run(tc.name, func(t *testing.T) {
            // ... (rest of your test logic, same as in previous examples) 
        })
    }
}

// ... in your TestParseLog function ...
{
    name: "Racial Spell Cast with Explore",
    simData: map[string]map[string]string{
        ExploreSheet: {
            fmt.Sprintf("S%d", 1): "1", // Explore is active
        },
        MagicSheet: {
            fmt.Sprintf("B%d", 1): "50",   // Mana available
            fmt.Sprintf("M%d", 1): "1",   // Some racial spell cast
        },
    },
    expected:     "Your wizards successfully cast Racial Spell at a cost of 150 mana.\n",
    expectedErr:  nil,
    simHour:     1,
},

func TestParseTechUnlock(t *testing.T) {
    testCases := []struct {
        name         string
        simData      map[string]map[string]string
        expected     string
        expectedErr  error
        simHour      int
    }{
        {
            name: "Tech Unlocked",
            simData: map[string]map[string]string{
                "Techs": {
                    fmt.Sprintf("K%d", 1): "1", // Tech unlocked
                    fmt.Sprintf("CA%d", 1): "Advanced Agriculture",
                },
            },
            expected:     "You have unlocked Advanced Agriculture\n",
            expectedErr:  nil,
            simHour:     1,
        },
        {
            name: "Tech Not Unlocked",
            simData: map[string]map[string]string{
                "Techs": {
                    fmt.Sprintf("K%d", 2): "0", // Tech not unlocked
                    fmt.Sprintf("CA%d", 2): "Super Steel", // This shouldn't be included in the output
                },
            },
            expected:     "",
            expectedErr:  nil,
            simHour:     2,
        },
        {
            name: "Tech Unlocked But Name Missing",
            simData: map[string]map[string]string{
                "Techs": {
                    fmt.Sprintf("K%d", 3): "1",   // Tech unlocked
                },
            }, // Missing tech name
            expected:     "",
            expectedErr:  fmt.Errorf("error reading tech name: %w", &excelize.ErrCellNotFound{}), // Expect a cell not found error
            simHour:     3,
        },
        {
            name: "Invalid Value for Tech Unlocked",
            simData: map[string]map[string]string{
                "Techs": {
                    fmt.Sprintf("K%d", 4): "invalid", // Invalid value
                    fmt.Sprintf("CA%d", 4): "Mysticism",
                },
            },
            expected:     "",
            expectedErr:  fmt.Errorf("error converting cell Techs!K4 value to int"), // Expect a conversion error
            simHour:     4,
        },
        // Add more test cases here for other error scenarios (e.g., missing sheet)
    }

    for _, tc := range testCases {
        t.Run(tc.name, func(t *testing.T) {
            mockSim := &SimMock{Data: tc.simData}
            glc := &GameLogCmd{sim: mockSim}

            result, err := glc.parseTechUnlock(tc.simHour)
            if err != nil && tc.expectedErr == nil {
                t.Errorf("Unexpected error: %v", err)
            } else if err == nil && tc.expectedErr != nil {
                t.Errorf("Expected error, but got none")
            } else if err != nil && tc.expectedErr != nil && !strings.Contains(err.Error(), tc.expectedErr.Error()) {
                t.Errorf("Incorrect error message: got %q, want substring %q", err, tc.expectedErr)
            }
            if result != tc.expected {
                t.Errorf("Incorrect result: got %q, want %q", result, tc.expected)
            }
        })
    }
}

func TestParseDailyPlatinum(t *testing.T) {
    testCases := []struct {
        name         string
        simData      map[string]map[string]string
        expected     string
        expectedErr  error
        simHour      int
    }{
        {
            name: "Platinum Awarded",
            simData: map[string]map[string]string{
                "Production": {
                    fmt.Sprintf("C%d", 1): "1", // Production is active
                },
                "Population": {
                    fmt.Sprintf("C%d", 1): "1000",
                },
            },
            expected:     "You have been awarded with 4000 platinum.\n",
            expectedErr:  nil,
            simHour:     1,
        },
        {
            name: "No Platinum Awarded (Production Zero)",
            simData: map[string]map[string]string{
                "Production": {
                    fmt.Sprintf("C%d", 2): "0", // No production
                },
                "Population": {
                    fmt.Sprintf("C%d", 2): "2000",
                },
            },
            expected:     "", // No message expected
            expectedErr:  nil,
            simHour:     2,
        },
        {
            name: "Error Reading Production Cell",
            simData:      map[string]map[string]string{}, // Missing Production data
            expected:     "",
            expectedErr:  fmt.Errorf("error reading cell Production!C3: %w", &excelize.ErrCellNotFound{}),
            simHour:     3,
        },
        {
            name: "Error Reading Population Cell",
            simData: map[string]map[string]string{
                "Production": {
                    fmt.Sprintf("C%d", 4): "1", // Production is active
                },
            }, // Missing Population data
            expected:     "",
            expectedErr:  fmt.Errorf("error reading cell Population!C4: %w", &excelize.ErrCellNotFound{}),
            simHour:     4,
        },
        // Add more test cases here for other error scenarios (e.g., invalid values)
    }

    // ... (rest of the test logic, similar to TestParseTechUnlock)
}

func TestParseNationalBank(t *testing.T) {
    testCases := []struct {
        name         string
        simData      map[string]map[string]string
        expected     string
        expectedErr  error
        simHour      int
    }{
        {
            name: "Trade Platinum for Lumber",
            simData: map[string]map[string]string{
                "Production": {
                    fmt.Sprintf("BC%d", 1): "-100",
                    fmt.Sprintf("BD%d", 1): "500",
                },
            },
            expected:     "100 platinum have been traded for 500 lumber.\n",
            expectedErr:  nil,
            simHour:     1,
        },
        {
            name: "Multiple Trades",
            simData: map[string]map[string]string{
                "Production": {
                    fmt.Sprintf("BC%d", 2): "250",
                    fmt.Sprintf("BD%d", 2): "-200",
                    fmt.Sprintf("BE%d", 2): "30",
                },
            },
            expected:     "200 lumber have been traded for 250 platinum and 30 ore.\n",
            expectedErr:  nil,
            simHour:     2,
        },
        // Add more test cases for other scenarios and error conditions
    }

    // ... (rest of the test logic, similar to TestParseTechUnlock)
}

func TestParseExplore(t *testing.T) {
    testCases := []struct {
        name         string
        simData      map[string]map[string]string
        expected     string
        expectedErr  error
        simHour      int
    }{
        {
            name: "Multiple Lands Explored",
            simData: map[string]map[string]string{
                "Explore": {
                    fmt.Sprintf("T%d", 1): "5",
                    fmt.Sprintf("V%d", 1): "2",
                    fmt.Sprintf("Z%d", 1): "3",
                    fmt.Sprintf("AH%d", 1): "1500",
                    fmt.Sprintf("AI%d", 1): "20",
                },
            },
            expected:     "Exploration for 5 Plains, 2 Mountains, 3 Water begun at a cost of 1500 platinum and 20 draftees.\n",
            expectedErr:  nil,
            simHour:     1,
        },
        {
            name: "No Exploration",
            simData: map[string]map[string]string{
                "Explore": {
                    fmt.Sprintf("T%d", 2): "0",
                    fmt.Sprintf("U%d", 2): "0",
                    fmt.Sprintf("V%d", 2): "0",
                    fmt.Sprintf("W%d", 2): "0",
                    fmt.Sprintf("X%d", 2): "0",
                    fmt.Sprintf("Y%d", 2): "0",
                    fmt.Sprintf("Z%d", 2): "0",
                },
            },
            expected:     "", // No exploration message
            expectedErr:  nil,
            simHour:     2,
        },
        // Add more test cases for error scenarios
        // ... (e.g., missing sheet, invalid cell values, errors reading cells)
    }

    // ... (test case execution logic, same as in previous examples)
}

func TestParseDailyLand(t *testing.T) {
    testCases := []struct {
        name         string
        simData      map[string]map[string]string
        expected     string
        expectedErr  error
        simHour      int
    }{
        {
            name: "Land Awarded",
            simData: map[string]map[string]string{
                "Explore": {
                    fmt.Sprintf("S%d", 1): "1", // Exploration is active
                },
                "Overview": {
                    "B70": "Plains",
                },
            },
            expected:     "You have been awarded with 20 Plains.\n",
            expectedErr:  nil,
            simHour:     1,
        },
        {
            name: "No Land Awarded (Exploration Inactive)",
            simData: map[string]map[string]string{
                "Explore": {
                    fmt.Sprintf("S%d", 2): "0", // Exploration is not active
                },
                "Overview": {
                    "B70": "Forest",
                },
            },
            expected:     "", // No message expected
            expectedErr:  nil,
            simHour:     2,
        },
        {
            name: "Error Reading Exploration Cell",
            simData:      map[string]map[string]string{}, // Missing Explore data
            expected:     "",
            expectedErr:  fmt.Errorf("error reading explore status: %w", &excelize.ErrCellNotFound{}),
            simHour:     3,
        },
        {
            name: "Error Reading Land Type Cell",
            simData: map[string]map[string]string{
                "Explore": {
                    fmt.Sprintf("S%d", 4): "1", // Exploration is active
                },
            }, // Missing Overview data
            expected:     "",
            expectedErr:  fmt.Errorf("error reading land type: %w", &excelize.ErrCellNotFound{}),
            simHour:     4,
        },
    }

    for _, tc := range testCases {
        t.Run(tc.name, func(t *testing.T) {
            mockSim := &SimMock{Data: tc.simData}
            glc := &GameLogCmd{sim: mockSim}

            result, err := glc.parseDailyLand(tc.simHour)
            if err != nil && tc.expectedErr == nil {
                t.Errorf("Unexpected error: %v", err)
            } else if err == nil && tc.expectedErr != nil {
                t.Errorf("Expected error, but got none")
            } else if err != nil && tc.expectedErr != nil && !strings.Contains(err.Error(), tc.expectedErr.Error()) {
                t.Errorf("Incorrect error message: got %q, want substring %q", err, tc.expectedErr)
            }
            if result != tc.expected {
                t.Errorf("Incorrect result: got %q, want %q", result, tc.expected)
            }
        })
    }
}

func TestParseDestruction(t *testing.T) {
    // ... your test cases similar to the previous examples ...
    testCases := []struct {
        name         string
        simData      map[string]map[string]string
        expected     string
        expectedErr  error
        simHour      int
    }{
        {
            name: "Multiple Buildings Destroyed",
            simData: map[string]map[string]string{
                "Construction": {
                    fmt.Sprintf("BW%d", 1): "2",    // Homes
                    fmt.Sprintf("BY%d", 1): "1",    // Smithies
                    fmt.Sprintf("CB%d", 1): "3",    // Lumber Yards
                },
            },
            expected:     "Destruction of 2 Homes, 1 Smithies, 3 Lumber Yards is complete.\n",
            expectedErr:  nil,
            simHour:     1,
        },
        {
            name: "No Destruction",
            simData: map[string]map[string]string{
                "Construction": {
                    // All building counts are 0
                },
            },
            expected:     "", // No message expected
            expectedErr:  nil,
            simHour:     2,
        },
        // Add more test cases for error scenarios
        // ... (e.g., missing sheet, invalid cell values, errors reading cells)
    }
    // ... your test logic
}

func TestParseRezone(t *testing.T) {
    testCases := []struct {
        name         string
        simData      map[string]map[string]string
        expected     string
        expectedErr  error
        simHour      int
    }{
        {
            name: "Multiple Lands Rezoned",
            simData: map[string]map[string]string{
                "Rezone": {
                    fmt.Sprintf("L%d", 1): "5",
                    fmt.Sprintf("N%d", 1): "-2",
                    fmt.Sprintf("R%d", 1): "1",
                    fmt.Sprintf("Y%d", 1): "500",
                },
            },
            expected:     "Rezoning begun at a cost of 500 platinum. The changes in land are as following: 5 Plains, -2 Mountains, 1 Water\n",
            expectedErr:  nil,
            simHour:     1,
        },
        {
            name: "No Rezoning",
            simData: map[string]map[string]string{
                "Rezone": {
                    // All rezoning counts are 0
                },
            },
            expected:     "", // No message expected
            expectedErr:  nil,
            simHour:     2,
        },
        // Add more test cases for error scenarios
        // ... (e.g., missing sheet, invalid cell values, errors reading cells)
    }

    // ... (test case execution logic, same as in previous examples)
}

func TestParseConstruction(t *testing.T) {
    testCases := []struct {
        name         string
        simData      map[string]map[string]string
        expected     string
        expectedErr  error
        simHour      int
    }{
        {
            name: "Multiple Buildings Constructed",
            simData: map[string]map[string]string{
                "Construction": {
                    fmt.Sprintf("O%d", 1):  "2",   // Homes
                    fmt.Sprintf("Q%d", 1):  "3",   // Farms
                    fmt.Sprintf("AA%d", 1): "1",   // Shrines
                    fmt.Sprintf("AQ%d", 1): "1500", // Platinum cost
                    fmt.Sprintf("AR%d", 1): "800",  // Lumber cost
                },
            },
            expected:     "Construction of 2 Homes, 3 Farms, 1 Shrines started at a cost of 1500 platinum and 800 lumber.\n",
            expectedErr:  nil,
            simHour:     1,
        },
        {
            name: "No Construction",
            simData: map[string]map[string]string{
                "Construction": {
                    // All building counts are 0
                },
            },
            expected:     "",
            expectedErr:  nil,
            simHour:     2,
        },
        // Add more test cases for error scenarios
        // ... (e.g., missing sheet, invalid cell values, errors reading cells)
    }

    // ... (test case execution logic, same as in previous examples)
}

func TestParseMilitaryTraining(t *testing.T) {
    testCases := []struct {
        name         string
        simData      map[string]map[string]string
        expected     string
        expectedErr  error
        simHour      int
    }{
        {
            name: "Multiple Units Trained",
            simData: map[string]map[string]string{
                "Military": {
                    "AG2": "Swordsmen", "AG1": "5",
                    "AH2": "Archers",   "AH1": "10",
                    "AL2": "Spies",     "AL1": "2",
                    "AN2": "Wizards",   "AN1": "1",
                    "AR1": "2000",
                    "AS1": "1500",
                },
            },
            expected:     "Training of 5 Swordsmen, 10 Archers, 2 Spies, 1 Wizards begun at a cost of 2000 platinum, 1500 ore, 15 draftees, 2 spies, and 1 wizards.\n",
            expectedErr:  nil,
            simHour:     1,
        },
        {
            name: "No Training",
            simData: map[string]map[string]string{
                "Military": {
                    // All training counts are 0
                },
            },
            expected:     "", // No message expected
            expectedErr:  nil,
            simHour:     2,
        },
        // Add more test cases for error scenarios
        // ... (e.g., missing sheet, invalid cell values, errors reading cells)
    }

    // ... (test case execution logic, same as in previous examples)
}

func TestParseImprovements(t *testing.T) {
    testCases := []struct {
        name         string
        simData      map[string]map[string]string
        expected     string
        expectedErr  error
        simHour      int
    }{
        {
            name: "Multiple Improvements",
            simData: map[string]map[string]string{
                "Imps": {
                    fmt.Sprintf("P%d", 1):  "100",
                    fmt.Sprintf("O%d", 1):  "platinum",
                    fmt.Sprintf("Q%d", 1):  "Trade",
                    fmt.Sprintf("S%d", 1):  "50",
                    fmt.Sprintf("R%d", 1):  "wood",
                    fmt.Sprintf("T%d", 1):  "Construction",
                },
            },
            expected:     "You invested 100 platinum into Trade.\nYou invested 50 wood into Construction.\n",
            expectedErr:  nil,
            simHour:     1,
        },
        {
            name: "No Improvements",
            simData: map[string]map[string]string{
                "Imps": {
                    fmt.Sprintf("P%d", 2): "0",
                    fmt.Sprintf("S%d", 2): "0",
                    fmt.Sprintf("V%d", 2): "0",
                },
            },
            expected:     "",
            expectedErr:  nil,
            simHour:     2,
        },
        // Add more test cases for error scenarios
        // ... (e.g., missing sheet, invalid cell values, errors reading cells)
    }

    // ... (test case execution logic, same as in previous examples)
}






