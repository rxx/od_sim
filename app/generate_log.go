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

const DEBUG = true

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
	if !DEBUG {
		return
	}

	pc, file, line, _ := runtime.Caller(1)
	funcName := runtime.FuncForPC(pc).Name()

	fmt.Printf("--- DEBUG on [%s:%s:%d] ---\n", filepath.Base(file), funcName, line)
	for _, value := range values {
		fmt.Println(value)
	}
	fmt.Println("--- ^_^ ---")
}

func executeGenerateLogCmd(path string) {
	sim, err := excelize.OpenFile(path)
	if err != nil {
		fmt.Println("Error on opening file %w", err)
		return
	}

	defer sim.Close()

	for hr := 0; hr <= LastHour; hr++ {
		action, err := generateWithHour(sim, hr)
		if err != nil {
			fmt.Println("Error on generating action", err)
			continue
		}

		fmt.Printf("Action on hour %v\n\n%v", hr+1, action)
	}
}

func currentHour(sim *excelize.File) (int, error) {
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

func generateWithHour(sim *excelize.File, hr int) (string, error) {
	var result strings.Builder

	// Starting at row 4 because of extra added row (due to uniform table headers)
	// simhr := hr + SimHr

	timeline, err := generateTimeline(sim, hr)
	if err != nil {
		return "", err
	}

	result.WriteString(timeline)

	return result.String(), nil
}

// TODO: Outputs
// ====== Protection Hour: 1  ( Local Time: 6:00:00 PM 5/18/2024 )  ( Domtime: 12:00:00 AM 5/18/2024 ) ======
// But seems correct ouput in next
// ====== Protection Hour: 1  ( Local Time: 6:00:00 PM 5/17/2024 )  ( Domtime: 12:00:00 AM 5/18/2024 ) ======
// Why 5/17?
func generateTimeline(sim *excelize.File, hr int) (string, error) {
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

	debugLog(localTimeValue, domTimeValue, dateValue)

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
			return "", fmt.Errorf("error parsing date: %w", err)
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
