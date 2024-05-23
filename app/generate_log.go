package main

import (
	"fmt"
	"path/filepath"
	"runtime"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

const (
	Overview     = "Overview"
	Population   = "Population"
	Production   = "Production"
	Construction = "Construction"
	Explore      = "Explore"
	Rezone       = "Rezone"
	Military     = "Military"
	Magic        = "Magic"
	Techs        = "Techs"
	Imps         = "Imps"
	Constants    = "Constants"
	Races        = "Races"
	LastHour     = 0 // 83
	SimHr        = 4
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

var sim *excelize.File

func executeGenerateLogCmd(path string) {
	var err error

	sim, err = excelize.OpenFile(path)
	if err != nil {
		fmt.Println("Error on opening file %w", err)
		return
	}

	defer sim.Close()

	for hr := 0; hr <= LastHour; hr++ {
		action, err := generateWithHour(hr)
		if err != nil {
			fmt.Println("Error on generating action", err)
			continue
		}

		fmt.Printf("Action on hour %v\n\n%v", hr+1, action)
	}
}

func currentHour() (int, error) {
	hourStr, err := sim.GetCellValue(Overview, "E17")
	if err != nil {
		return 0, fmt.Errorf("read from sim: %w", err)
	}

	hour, err := strconv.Atoi(hourStr)
	if err != nil {
		return 0, fmt.Errorf("converting current hour: %w", err)
	}

	hour -= 1
	if hour < 0 {
		hour = 0
	}

	return hour, nil

	// SimHour hournum, "msg"
}

func generateWithHour(hr int) (string, error) {
	var sb strings.Builder

	// Starting at row 4 because of extra added row (due to uniform table headers)
	simhr := hr + SimHr

	timeline, err := generateTimeline(hr)
	if err != nil {
		return "", err
	}
	sb.WriteString(timeline)
	sb.WriteString("\n")

	draftrate, err := setDraftRate(simhr)
	if err != nil {
		return "", err
	}
	sb.WriteString(draftrate)
	sb.WriteString("\n")

	sb.WriteString(releaseUnits(simhr))
	sb.WriteString("\n")

	return sb.String(), nil
}

// TODO: Outputs
// ====== Protection Hour: 1  ( Local Time: 6:00:00 PM 5/18/2024 )  ( Domtime: 12:00:00 AM 5/18/2024 ) ======
// But seems correct ouput in next
// ====== Protection Hour: 1  ( Local Time: 6:00:00 PM 5/17/2024 )  ( Domtime: 12:00:00 AM 5/18/2024 ) ======
// Why 5/17?
func generateTimeline(hr int) (string, error) {
	localTimeCell := fmt.Sprintf("BY%d", hr+SimHr)
	domTimeCell := fmt.Sprintf("BZ%d", hr+SimHr)

	localTimeValue, err := sim.GetCellValue(Imps, localTimeCell)
	if err != nil {
		return "", fmt.Errorf("error reading local time: %w", err)
	}

	domTimeValue, err := sim.GetCellValue(Imps, domTimeCell)
	if err != nil {
		return "", fmt.Errorf("error reading dom time: %w", err)
	}

	dateValue, err := sim.GetCellValue(Overview, "B15")
	if err != nil {
		return "", fmt.Errorf("error reading date: %w", err)
	}

	debugLog("LocalTime", localTimeValue, "DomTime", domTimeValue, "date", dateValue)

	localTime, err := time.Parse("15:04", localTimeValue)
	if err != nil {
		return "", fmt.Errorf("error parsing local time: %w", err)
	}

	domTime, err := time.Parse("15:04", domTimeValue)
	if err != nil {
		return "", fmt.Errorf("error parsing dom time: %w", err)
	}

	date, err := time.Parse("1/2/2006", dateValue)
	if err != nil {
		date, err = time.Parse("1-2-06", dateValue)
		if err != nil {
			date, err = time.Parse("2006/01/02", dateValue)
			if err != nil {
				return "", fmt.Errorf("error parsing date: %w", err)
			}
		}
	}

	localTime = time.Date(date.Year(), date.Month(), date.Day(),
		localTime.Hour(), localTime.Minute(), 0, 0, time.UTC)

	domTime = time.Date(date.Year(), date.Month(), date.Day(),
		domTime.Hour(), domTime.Minute(), 0, 0, time.UTC)

	localTimeLong := localTime.Format("3:04:05 PM")
	localTimeShort := localTime.Format("1/2/2006")
	domTimeLong := domTime.Format("3:04:05 PM")
	domTimeShort := domTime.Format("1/2/2006")

	var timeline strings.Builder
	timeline.WriteString("====== Protection Hour: ")
	timeline.WriteString(fmt.Sprintf("%d", hr+1))
	timeline.WriteString("  ( Local Time: ")
	timeline.WriteString(localTimeLong)
	timeline.WriteString(" ")
	timeline.WriteString(localTimeShort)
	timeline.WriteString(" )  ( Domtime: ")
	timeline.WriteString(domTimeLong)
	timeline.WriteString(" ")
	timeline.WriteString(domTimeShort)
	timeline.WriteString(" ) ======")

	return timeline.String(), nil
}

func setDraftRate(simhr int) (string, error) {
	currentRateCell := fmt.Sprintf("Y%d", simhr)
	previousRateCell := fmt.Sprintf("Z%d", simhr-1)

	currentRateStr, err := sim.GetCellValue(Military, currentRateCell)
	if err != nil {
		return "", fmt.Errorf("error reading current draftrate: %w", err)
	}

	previousRateStr, err := sim.GetCellValue(Military, previousRateCell)
	if err != nil {
		return "", fmt.Errorf("error reading previous draftrate: %w", err)
	}

	debugLog("CurrentDraftrate", currentRateStr, "PreviousDraftrate", previousRateStr)

	var buf strings.Builder

	if currentRateStr == "" || currentRateStr == previousRateStr {
		return "", nil
	}

	buf.WriteString("Draftrate changed to ")
	buf.WriteString(currentRateStr)
	buf.WriteString(".")

	return buf.String(), nil
}

func releaseUnits(simhr int) string {
	// Read unit names and unit counts
	unitNames := make([]string, 8)
	units := make([]int, 8)
	cols := []string{"AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE"}
	for i, col := range cols {
		// Read unit names from row 2
		unitNameCell := fmt.Sprintf("%s2", col)
		unitNames[i], _ = sim.GetCellValue(Military, unitNameCell)

		// Read unit counts from simhr row
		unitCountCell := fmt.Sprintf("%s%d", col, simhr)
		unitCountStr, _ := sim.GetCellValue(Military, unitCountCell)
		units[i], _ = strconv.Atoi(unitCountStr) // Parse to int (assuming integer values)
	}

	// Read draftees count from AW column
	drafteesCell := fmt.Sprintf("AW%d", simhr)
	drafteesStr, _ := sim.GetCellValue(Military, drafteesCell)
	draftees, _ := strconv.Atoi(drafteesStr) // Parse to int

	// Check for Released Units and Build Message
	var sb strings.Builder
	released := false
	addedUnits := 0

	sb.WriteString("You successfully released ")

	for i := 0; i < len(units); i++ {
		if units[i] > 0 {
			released = true
			if addedUnits > 0 {
				sb.WriteString(", ") // Add comma before each unit except the first
			}
			addedUnits++
			sb.WriteString(fmt.Sprintf("%d %s", units[i], unitNames[i]))
		}
	}

	if released {
		sb.WriteString(".\n")
	} else {
		sb.Reset()
	}

	// 3. Check for Draftees
	if draftees > 0 {
		sb.WriteString(fmt.Sprintf("You successfully released %d draftees into the peasantry.\n", draftees))
	}

	return sb.String()
}
