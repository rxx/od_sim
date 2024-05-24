package main

import (
	"fmt"
	"runtime/debug"
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
	LastHour     = 72
)

type ActionFunc func() (string, error)

type Sim interface {
	GetCellValue(sheet, cell string, opts ...excelize.Options) (string, error)
	Close() error
}

type GameLogCmd struct {
	currentHour int
	simHour     int
	simPath     string
	sim         Sim
	// sim     *excelize.File
	output  strings.Builder
	actions []ActionFunc
}

func NewGameLog(path string) *GameLogCmd {
	gameLogCmd := &GameLogCmd{
		simPath:     path,
		currentHour: 0,
	}
	gameLogCmd.initActions()

	return gameLogCmd
}

func (c *GameLogCmd) initActions() {
	c.actions = append(c.actions,
		c.tickAction,
		c.draftRateAction,
		c.releaseUnitsAction)
}

// Starting at row 4 because of extra added row (due to uniform table headers)
func (c *GameLogCmd) setCurrentHour(hr int) {
	c.currentHour = hr
	c.simHour = hr + 4
}

func (c *GameLogCmd) Execute() {
	var err error

	c.sim, err = excelize.OpenFile(c.simPath)
	if err != nil {
		fmt.Println("Error on opening file %w", err)
		return
	}

	defer c.sim.Close()

	// for hr := 0; hr <= LastHour; hr++ {
	// c.setCurrentHour(hr)
	c.setCurrentHour(0)

	for _, actionFunc := range c.actions {
		result, err := actionFunc()
		if err != nil {
			c.output.WriteString(fmt.Sprintf("Error on executing action: %v", err))
			c.output.WriteString("\n")

			if debugEnabled {
				debug.PrintStack()
			}
			break
		}

		if result != "" {
			c.output.WriteString(result)
			c.output.WriteString("\n")
		}
	}

	fmt.Println(c.output.String())
}

func (c *GameLogCmd) tickAction() (string, error) {
	localTimeCell := fmt.Sprintf("BY%d", c.simHour)
	domTimeCell := fmt.Sprintf("BZ%d", c.simHour)

	localTimeValue, err := c.sim.GetCellValue(Imps, localTimeCell)
	if err != nil {
		return "", fmt.Errorf("error reading local time: %w", err)
	}

	domTimeValue, err := c.sim.GetCellValue(Imps, domTimeCell)
	if err != nil {
		return "", fmt.Errorf("error reading dom time: %w", err)
	}

	dateValue, err := c.sim.GetCellValue(Overview, "B15")
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
	timeline.WriteString(fmt.Sprintf("%d", c.currentHour+1))
	timeline.WriteString(" ( Local Time: ")
	timeline.WriteString(localTimeLong)
	timeline.WriteString(" ")
	timeline.WriteString(localTimeShort)
	timeline.WriteString(" ) ( Domtime: ")
	timeline.WriteString(domTimeLong)
	timeline.WriteString(" ")
	timeline.WriteString(domTimeShort)
	timeline.WriteString(" ) ======")

	return timeline.String(), nil
}

func (c *GameLogCmd) draftRateAction() (string, error) {
	currentRateCell := fmt.Sprintf("Y%d", c.simHour)
	previousRateCell := fmt.Sprintf("Z%d", c.simHour-1)

	currentRateStr, err := c.sim.GetCellValue(Military, currentRateCell)
	if err != nil {
		return "", fmt.Errorf("error reading current draftrate: %w", err)
	}

	previousRateStr, err := c.sim.GetCellValue(Military, previousRateCell)
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

	return buf.String(), nil
}

func (c *GameLogCmd) releaseUnitsAction() (string, error) {
	// Read unit names and unit counts
	unitNames := make([]string, 8)
	units := make([]int, 8)
	cols := []string{"AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE"}
	for i, col := range cols {
		// Read unit names from row 2
		unitNameCell := fmt.Sprintf("%s2", col)
		unitNames[i], _ = c.sim.GetCellValue(Military, unitNameCell)

		// Read unit counts from simhr row
		unitCountCell := fmt.Sprintf("%s%d", col, c.simHour)
		unitCountStr, _ := c.sim.GetCellValue(Military, unitCountCell)
		units[i], _ = strconv.Atoi(unitCountStr) // Parse to int (assuming integer values)
	}

	// Read draftees count from AW column
	drafteesCell := fmt.Sprintf("AW%d", c.simHour)
	drafteesStr, _ := c.sim.GetCellValue(Military, drafteesCell)
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
		sb.WriteString("\n")
	} else {
		sb.Reset()
	}

	// 3. Check for Draftees
	if draftees > 0 {
		sb.WriteString(fmt.Sprintf("You successfully released %d draftees into the peasantry\n", draftees))
	}

	return sb.String(), nil
}
