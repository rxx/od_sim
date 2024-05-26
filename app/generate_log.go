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
  
  // Magic
  GaiasWatch     = "Gaia's Watch"
    MiningStrength = "Mining Strength"
    AresCall       = "Ares' Call"
    MidasTouch     = "Midas Touch"
    Harmony        = "Harmony"
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
	gameLogCmd.initSim()

	return gameLogCmd
}

func (c *GameLogCmd) initActions() {
	c.actions = []ActionFunc{
		c.tickAction,
		c.draftRateAction,
		c.releaseUnitsAction,
	}
}

// Starting at row 4 because of extra added row (due to uniform table headers)
func (c *GameLogCmd) setCurrentHour(hr int) {
	c.currentHour = hr
	c.simHour = hr + 4
}

func (c *GameLogCmd) initSim() {
	var err error

	c.sim, err = excelize.OpenFile(c.simPath)
	if err != nil {
		fmt.Println("Error on opening file %w", err)
		return
	}
}

func (c *GameLogCmd) readValue(sheet, cell, errorMsg string) (string, error) {
	value, err := c.sim.GetCellValue(sheet, cell)
	if err != nil {
		return "", fmt.Errorf("%s: %w", errorMsg, err)
	}

	return value, nil
}

func (c *GameLogCmd) Execute() {
	defer c.sim.Close()

	// for hr := 0; hr <= LastHour; hr++ {
	// c.setCurrentHour(hr)
	// }
	c.setCurrentHour(0)
	c.executeActions()

	fmt.Println(c.output.String())
}

func (c *GameLogCmd) executeActions() {
	for _, actionFunc := range c.actions {
		result, err := actionFunc()
		if err != nil {
			c.output.WriteString(fmt.Sprintf("Error on executing action: %v", err))
			c.output.WriteString("\n")

			if cmdVars.debugEnabled {
				debug.PrintStack()
			}
			break
		}

		if result != "" {
			c.output.WriteString(result)
			c.output.WriteString("\n")
		}
	}
}

func (c *GameLogCmd) tickAction() (string, error) {
	const dateCell = "B15"
	localTimeCell := fmt.Sprintf("BY%d", c.simHour)
	domTimeCell := fmt.Sprintf("BZ%d", c.simHour)

	localTimeValue, err := c.readValue(Imps, localTimeCell, "error reading local time")
	if err != nil {
		return "", err
	}

	domTimeValue, err := c.readValue(Imps, domTimeCell, "error reading dom time")
	if err != nil {
		return "", err
	}

	dateValue, err := c.readValue(Overview, dateCell, "error reading date")
	if err != nil {
		return "", err
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

	currentRateStr, err := c.readValue(Military, currentRateCell, "error reading current draftrate")
	if err != nil {
		return "", err
	}

	previousRateStr, err := c.readValue(Military, previousRateCell, "error reading previous draftrate")
	if err != nil {
		return "", err
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
	var err error

	// Read unit names and unit counts
	unitNames := make([]string, 8)
	units := make([]int, 8)
	cols := []string{"AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE"}
	for i, col := range cols {
		// Read unit names from row 2
		unitNameCell := fmt.Sprintf("%s2", col)
		unitNames[i], err = c.readValue(Military, unitNameCell, "error reading unit name")
		if err != nil {
			return "", err
		}

		// Read unit counts from simhr row
		unitCountCell := fmt.Sprintf("%s%d", col, c.simHour)
		unitCountStr, err := c.readValue(Military, unitCountCell, "error reading unit value")
		if err != nil {
			return "", err
		}

		if unitCountStr == "" {
			continue
		}

		units[i], err = strconv.Atoi(unitCountStr)
		if err != nil {
			return "", fmt.Errorf("error parsing unit value: %w", err)
		}
	}

	// Read draftees count from AW column
	drafteesCell := fmt.Sprintf("AW%d", c.simHour)
	drafteesStr, err := c.readValue(Military, drafteesCell, "error reading draftees value")
	if err != nil {
		return "", err
	}

	draftees := 0
	if drafteesStr != "" {
		draftees, err = strconv.Atoi(drafteesStr)
		if err != nil {
			return "", fmt.Errorf("error parsing draftees value: %w", err)
		}
	}

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

	if draftees > 0 {
		sb.WriteString(fmt.Sprintf("You successfully released %d draftees into the peasantry\n", draftees))
	}

	return sb.String(), nil
}

_______

func (c *GameLogCmd) parseMagicActions(simHour int) (string, error) {
    var actions strings.Builder

    getIntValue := func(sheet, axis string) (int, error) {
        val, err := c.sim.GetCellValue(sheet, axis)
        if err != nil {
            return 0, fmt.Errorf("error reading cell %s!%s: %w", sheet, axis, err)
        }
        intVal, err := strconv.Atoi(val)
        if err != nil {
            return 0, fmt.Errorf("error converting cell %s!%s value to int: %w", sheet, axis, err)
        }
        return intVal, nil
    }

    checkAndCastSpell := func(spellName, exploreCell, magicCell string, costMultiplier float64) error {
        exploreVal, err := getIntValue(ExploreSheet, exploreCell)
        if err != nil {
            return err 
        }
        magicVal, err := getIntValue(MagicSheet, magicCell)
        if err != nil {
            return err
        }

        mana := 0
        if exploreVal != 0 && magicVal != 0 {
            mana = (magicVal - 20) * 2
        } else if magicVal != 0 {
            mana = magicVal * 2
        } else {
            return nil // No spell was cast, so no message to add
        }
        actions.WriteString(fmt.Sprintf("Your wizards successfully cast %s at a cost of %d mana.\n", spellName, int(float64(mana)*costMultiplier)))
        return nil
    }

    // Check and cast each spell
    for _, spell := range []struct {
        name         string
        exploreCell string
        magicCell  string
        costMultiplier float64
    }{
        {GaiasWatch, fmt.Sprintf("S%d", simHour), fmt.Sprintf("G%d", simHour), 1.0},
        {MiningStrength, fmt.Sprintf("S%d", simHour), fmt.Sprintf("H%d", simHour), 1.0},
        {AresCall, fmt.Sprintf("S%d", simHour), fmt.Sprintf("I%d", simHour), 2.5},
        {MidasTouch, fmt.Sprintf("S%d", simHour), fmt.Sprintf("J%d", simHour), 2.5},
        {Harmony, fmt.Sprintf("S%d", simHour), fmt.Sprintf("K%d", simHour), 2.5},
    } {
        if err := checkAndCastSpell(spell.name, spell.exploreCell, spell.magicCell, spell.costMultiplier); err != nil {
            return "", err // Propagate any errors
        }
    }
    return actions.String(), nil
}

package main

import (
    "fmt"
    "strconv"
    "strings"
)

// ... (your other types, constants, and helper functions)

const (
    RacialSpell = "Racial Spell" 
    // ... other constants
)
 
func (c *GameLogCmd) parseRacialSpell(simHour int) (string, error) {
    var actions strings.Builder

    // Helper function to read a cell value and convert to int (reuse from previous code)
    getIntValue := func(sheet, axis string) (int, error) {
        // ... (same implementation as before)
    }

    // Check if any racial spell is cast
    racialSpellCast := false
    for col := 'L'; col <= 'U'; col++ {
        cell := fmt.Sprintf("%c%d", col, simHour)
        val, err := getIntValue(MagicSheet, cell)
        if err != nil {
            return "", err
        }
        if val != 0 {
            racialSpellCast = true
            break
        }
    }

    if racialSpellCast {
        exploreVal, err := getIntValue(ExploreSheet, fmt.Sprintf("S%d", simHour))
        if err != nil {
            return "", err
        }
        manaVal, err := getIntValue(MagicSheet, fmt.Sprintf("B%d", simHour))
        if err != nil {
            return "", err
        }
        
        mana := manaVal * 5
        if exploreVal != 0 {
            mana = (manaVal - 20) * 5
        }

        actions.WriteString(fmt.Sprintf("Your wizards successfully cast %s at a cost of %d mana.\n", RacialSpell, mana))
    }

    return actions.String(), nil
}

func (c *GameLogCmd) parseTechUnlock(simHour int) (string, error) {
    // Check if a tech was unlocked
    techUnlocked, err := c.sim.GetCellValue("Techs", fmt.Sprintf("K%d", simHour))
    if err != nil {
        return "", fmt.Errorf("error reading tech unlocked status: %w", err)
    }
    
    if techUnlocked != "0" { // Check if tech was unlocked
        techName, err := c.sim.GetCellValue("Techs", fmt.Sprintf("CA%d", simHour))
        if err != nil {
            return "", fmt.Errorf("error reading tech name: %w", err)
        }

        return fmt.Sprintf("You have unlocked %s\n", techName), nil
    }

    return "", nil // No tech unlocked
}

package main

import (
    "fmt"
    "strconv"
    "strings"
)

// ... (your other types, constants, etc.)
const PLATINUM_ACTION = "daily_platinum" // Added a constant for the action type

func (c *GameLogCmd) parseDailyPlatinum(simHour int) (string, error) {
    var actions strings.Builder

    // Helper function to get a cell value as an integer (similar to before)
    getIntValue := func(sheet, axis string) (int, error) {
        val, err := c.sim.GetCellValue(sheet, axis)
        if err != nil {
            return 0, fmt.Errorf("error reading cell %s!%s: %w", sheet, axis, err)
        }
        intVal, err := strconv.Atoi(val)
        if err != nil {
            return 0, fmt.Errorf("error converting cell %s!%s value to int: %w", sheet, axis, err)
        }
        return intVal, nil
    }

    productionValue, err := getIntValue("Production", fmt.Sprintf("C%d", simHour))
    if err != nil {
        return "", err // Propagate the error
    }

    if productionValue != 0 {
        populationValue, err := getIntValue("Population", fmt.Sprintf("C%d", simHour))
        if err != nil {
            return "", err // Propagate the error
        }

        platinumAwarded := populationValue * 4
        actions.WriteString(fmt.Sprintf("You have been awarded with %d platinum.\n", platinumAwarded))
    }

    return actions.String(), nil
}

package main

import (
    "fmt"
    "strconv"
    "strings"
)

// ... (other types, constants, etc.)
const BANK_ACTION = "national_bank"

func (c *GameLogCmd) parseNationalBank(simHour int) (string, error) {
    var actions strings.Builder

    // Helper function to get a cell value as an integer
    getIntValue := func(sheet, axis string) (int, error) {
        val, err := c.sim.GetCellValue(sheet, axis)
        if err != nil {
            return 0, fmt.Errorf("error reading cell %s!%s: %w", sheet, axis, err)
        }
        intVal, err := strconv.Atoi(val)
        if err != nil {
            return 0, fmt.Errorf("error converting cell %s!%s value to int: %w", sheet, axis, err)
        }
        return intVal, nil
    }

    // Read values from "Production" sheet
    plat, _ := getIntValue("Production", fmt.Sprintf("BC%d", simHour))
    lumber, _ := getIntValue("Production", fmt.Sprintf("BD%d", simHour))
    ore, _ := getIntValue("Production", fmt.Sprintf("BE%d", simHour))
    gems, _ := getIntValue("Production", fmt.Sprintf("BF%d", simHour))

    if plat != 0 || lumber != 0 || ore != 0 || gems != 0 { // Check if any exchange happened
        var tradedItems []string
        var receivedItems []string

        // Build traded items string
        addItem := func(item string, amount int) {
            if amount < 0 {
                tradedItems = append(tradedItems, fmt.Sprintf("%d %s", -amount, item))
            } else if amount > 0 {
                receivedItems = append(receivedItems, fmt.Sprintf("%d %s", amount, item))
            }
        }
        addItem("platinum", plat)
        addItem("lumber", lumber)
        addItem("ore", ore)
        addItem("gems", gems)

        // Construct the action message
        if len(tradedItems) > 0 {
            actions.WriteString(strings.Join(tradedItems, ", ") + " have been traded for ")
        }
        if len(receivedItems) > 0 {
            actions.WriteString(strings.Join(receivedItems, " and ") + ".\n")
        }
    }

    return actions.String(), nil
}

package main

import (
    "fmt"
    "strconv"
    "strings"
)

// ... (other types, constants, etc.)
const EXPLORE_ACTION = "explore"

func (c *GameLogCmd) parseExplore(simHour int) (string, error) {
    var actions strings.Builder
    getIntValue := func(sheet, axis string) (int, error) {
        // ... (same implementation as before)
    }
    landTypes := []string{"Plains", "Forest", "Mountains", "Hills", "Swamps", "Caverns", "Water"}
    exploreCounts := make([]int, len(landTypes))

    // Read exploration counts for each land type
    for i, landType := range landTypes {
        cellAxis := fmt.Sprintf("%c%d", 'T'+i, simHour) // Calculate cell addresses (T, U, V, ...)
        count, err := getIntValue(ExploreSheet, cellAxis)
        if err != nil {
            return "", err
        }
        exploreCounts[i] = count
    }

    // Check if any exploration happened
    exploreOccurred := false
    for _, count := range exploreCounts {
        if count > 0 {
            exploreOccurred = true
            break
        }
    }

    if exploreOccurred {
        actions.WriteString("Exploration for ")

        // Build list of explored lands
        for i, count := range exploreCounts {
            if count > 0 {
                if actions.Len() > len("Exploration for ") { // Add comma if not the first item
                    actions.WriteString(", ")
                }
                actions.WriteString(fmt.Sprintf("%d %s", count, landTypes[i]))
            }
        }

        // Read cost values
        platCost, _ := getIntValue(ExploreSheet, fmt.Sprintf("AH%d", simHour))
        drafteeCost, _ := getIntValue(ExploreSheet, fmt.Sprintf("AI%d", simHour))

        actions.WriteString(fmt.Sprintf(" begun at a cost of %d platinum and %d draftees.\n", platCost, drafteeCost))
    }

    return actions.String(), nil
}

package main

import (
    "fmt"
    "strings"
)

// ... (other types, constants, etc.)
const DAILY_LAND_ACTION = "daily_land"

func (c *GameLogCmd) parseDailyLand(simHour int) (string, error) {
    var actions strings.Builder

    // Check if exploration is active
    exploreVal, err := c.sim.GetCellValue("Explore", fmt.Sprintf("S%d", simHour))
    if err != nil {
        return "", fmt.Errorf("error reading explore status: %w", err)
    }

    if exploreVal != "0" {
        landType, err := c.sim.GetCellValue("Overview", "B70")
        if err != nil {
            return "", fmt.Errorf("error reading land type: %w", err)
        }

        actions.WriteString(fmt.Sprintf("You have been awarded with 20 %s.\n", landType))
    }

    return actions.String(), nil
}

package main

import (
    "fmt"
    "strconv"
    "strings"
)

// ... other constants and types

const DESTRUCTION_ACTION = "destruction"

func (c *GameLogCmd) parseDestruction(simHour int) (string, error) {
    var actions strings.Builder

    // Helper function to get a cell value as an integer
    getIntValue := func(axis string) (int, error) {
        val, err := c.sim.GetCellValue("Construction", axis)
        if err != nil {
            return 0, fmt.Errorf("error reading cell Construction!%s: %w", axis, err)
        }
        intVal, err := strconv.Atoi(val)
        if err != nil {
            return 0, fmt.Errorf("error converting cell Construction!%s value to int: %w", axis, err)
        }
        return intVal, nil
    }

    // Read building destruction counts
    buildingCounts := make([]int, 18) // 18 building types (including Homes)
    buildingNames := []string{
        "Homes", "Alchemies", "Farms", "Smithies", "Masonries", "Lumber Yards", "Forest Havens",
        "Ore Mines", "Gryphon Nests", "Factories", "Guard Towers", "Barracks", "Shrines", "Towers",
        "Temples", "Wizard Guilds", "Diamond Mines", "Schools", "Docks",
    }
    cols := []string{"BW", "BX", "BY", "BZ", "CA", "CB", "CC", "CD", "CE", "CF", "CG", "CH", "CI", "CJ", "CK", "CL", "CM", "CN", "CO"}

    for i, col := range cols {
        buildingCounts[i], _ = getIntValue(fmt.Sprintf("%s%d", col, simHour))
    }

    // Check if any destruction occurred
    destructionOccurred := false
    for _, count := range buildingCounts {
        if count > 0 {
            destructionOccurred = true
            break
        }
    }

    if destructionOccurred {
        actions.WriteString("Destruction of ")
        comma := false

        // Add destroyed buildings to message
        for i, count := range buildingCounts {
            if count > 0 {
                if comma {
                    actions.WriteString(", ")
                }
                actions.WriteString(fmt.Sprintf("%d %s", count, buildingNames[i]))
                comma = true
            }
        }

        actions.WriteString(" is complete.\n")
    }

    return actions.String(), nil
}

package main

import (
    "fmt"
    "strconv"
    "strings"
)

// ... (other types and constants)
const REZONE_ACTION = "rezone"

func (c *GameLogCmd) parseRezone(simHour int) (string, error) {
    var actions strings.Builder

    getIntValue := func(axis string) (int, error) {
        val, err := c.sim.GetCellValue("Rezone", axis)
        if err != nil {
            return 0, fmt.Errorf("error reading cell Rezone!%s: %w", axis, err)
        }
        intVal, err := strconv.Atoi(val)
        if err != nil {
            return 0, fmt.Errorf("error converting cell Rezone!%s value to int: %w", axis, err)
        }
        return intVal, nil
    }

    // Read rezoning counts for each land type
    landTypes := []string{"Plains", "Forest", "Mountains", "Hills", "Swamps", "Caverns", "Water"}
    rezoneCounts := make([]int, len(landTypes))
    for i, landType := range landTypes {
        cellAxis := fmt.Sprintf("%c%d", 'L'+i, simHour)
        rezoneCounts[i], _ = getIntValue(cellAxis)
    }

    // Check if any rezoning occurred
    rezoningOccurred := false
    for _, count := range rezoneCounts {
        if count != 0 {
            rezoningOccurred = true
            break
        }
    }

    if rezoningOccurred {
        // Read cost
        platCost, _ := getIntValue(fmt.Sprintf("Y%d", simHour))
        // Build message
        actions.WriteString(fmt.Sprintf("Rezoning begun at a cost of %d platinum. The changes in land are as following: ", platCost))

        comma := false
        for i, count := range rezoneCounts {
            if count != 0 {
                if comma {
                    actions.WriteString(", ")
                }
                actions.WriteString(fmt.Sprintf("%d %s", count, landTypes[i]))
                comma = true
            }
        }
        actions.WriteString("\n")
    }

    return actions.String(), nil
}

package main

import (
    "fmt"
    "strconv"
    "strings"
)

// ... (other types and constants)
const CONSTRUCTION_ACTION = "construction"

func (c *GameLogCmd) parseConstruction(simHour int) (string, error) {
    var actions strings.Builder

    getIntValue := func(axis string) (int, error) {
        val, err := c.sim.GetCellValue("Construction", axis)
        if err != nil {
            return 0, fmt.Errorf("error reading cell Construction!%s: %w", axis, err)
        }
        intVal, err := strconv.Atoi(val)
        if err != nil {
            return 0, fmt.Errorf("error converting cell Construction!%s value to int: %w", axis, err)
        }
        return intVal, nil
    }

    // Read building construction counts
    buildingCounts := make([]int, 18) // 18 building types (including Homes)
    buildingNames := []string{
        "Homes", "Alchemies", "Farms", "Smithies", "Masonries", "Lumber Yards", "Forest Havens",
        "Ore Mines", "Gryphon Nests", "Factories", "Guard Towers", "Barracks", "Shrines", "Towers",
        "Temples", "Wizard Guilds", "Diamond Mines", "Schools", "Docks",
    }
    cols := []string{"O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG"}

    for i, col := range cols {
        buildingCounts[i], _ = getIntValue(fmt.Sprintf("%s%d", col, simHour))
    }

    // Check if any construction occurred
    constructionOccurred := false
    for _, count := range buildingCounts {
        if count > 0 {
            constructionOccurred = true
            break
        }
    }

    if constructionOccurred {
        actions.WriteString("Construction of ")
        comma := false

        // Add constructed buildings to message
        for i, count := range buildingCounts {
            if count > 0 {
                if comma {
                    actions.WriteString(", ")
                }
                actions.WriteString(fmt.Sprintf("%d %s", count, buildingNames[i]))
                comma = true
            }
        }

        // Read cost values
        platCost, _ := getIntValue(fmt.Sprintf("AQ%d", simHour))
        lumberCost, _ := getIntValue(fmt.Sprintf("AR%d", simHour))

        actions.WriteString(fmt.Sprintf(" started at a cost of %d platinum and %d lumber.\n", platCost, lumberCost))
    }

    return actions.String(), nil
}

package main

import (
    "fmt"
    "strconv"
    "strings"
)

// ... (other types and constants)
const MILITARY_TRAINING_ACTION = "military_training"

func (c *GameLogCmd) parseMilitaryTraining(simHour int) (string, error) {
    var actions strings.Builder

    getIntValue := func(axis string) (int, error) {
        val, err := c.sim.GetCellValue("Military", axis)
        if err != nil {
            return 0, fmt.Errorf("error reading cell Military!%s: %w", axis, err)
        }
        intVal, err := strconv.Atoi(val)
        if err != nil {
            return 0, fmt.Errorf("error converting cell Military!%s value to int: %w", axis, err)
        }
        return intVal, nil
    }

    // Read unit training counts and names
    unitNames := make([]string, 8)
    unitCounts := make([]int, 8)
    cols := []string{"AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN"}
    for i, col := range cols {
        // Read unit names from row 2
        unitNameCell := fmt.Sprintf("%s2", col)
        unitNames[i], _ = c.sim.GetCellValue("Military", unitNameCell)

        // Read unit counts from simhr row
        unitCountCell := fmt.Sprintf("%s%d", col, simHour)
        unitCounts[i], _ = getIntValue(unitCountCell)
    }

    // Calculate total draftees used (excluding spies and wizards)
    draftees := 0
    for i := 0; i < 5; i++ {
        draftees += unitCounts[i]
    }

    // Check if any training occurred
    trainingOccurred := false
    for _, count := range unitCounts {
        if count > 0 {
            trainingOccurred = true
            break
        }
    }

    if trainingOccurred {
        actions.WriteString("Training of ")
        comma := false
        for i, count := range unitCounts {
            if count > 0 {
                if comma {
                    actions.WriteString(", ")
                }
                actions.WriteString(fmt.Sprintf("%d %s", count, unitNames[i]))
                comma = true
            }
        }

        // Read cost values
        platCost, _ := getIntValue(fmt.Sprintf("AR%d", simHour))
        oreCost, _ := getIntValue(fmt.Sprintf("AS%d", simHour))
        spyCount := unitCounts[5] // Index 5 corresponds to spies
        wizardCount := unitCounts[7] // Index 7 corresponds to wizards

        actions.WriteString(fmt.Sprintf(" begun at a cost of %d platinum, %d ore, %d draftees, %d spies, and %d wizards.\n", platCost, oreCost, draftees, spyCount, wizardCount))
    }

    return actions.String(), nil
}

package main

import (
    "fmt"
    "strconv"
    "strings"
)

// ... (other types and constants)
const IMPROVEMENT_ACTION = "improvement"

func (c *GameLogCmd) parseImprovements(simHour int) (string, error) {
    var actions strings.Builder

    // Helper function to get a cell value as a string
    getStringValue := func(axis string) (string, error) {
        val, err := c.sim.GetCellValue("Imps", axis)
        if err != nil {
            return "", fmt.Errorf("error reading cell Imps!%s: %w", axis, err)
        }
        return val, nil
    }

    // Helper function to get a cell value as an integer
    getIntValue := func(axis string) (int, error) {
        val, err := c.sim.GetCellValue("Imps", axis)
        if err != nil {
            return 0, fmt.Errorf("error reading cell Imps!%s: %w", axis, err)
        }
        intVal, err := strconv.Atoi(val)
        if err != nil {
            return 0, fmt.Errorf("error converting cell Imps!%s value to int: %w", axis, err)
        }
        return intVal, nil
    }

    // Function to check for an improvement and format the action message
    checkAndFormatImprovement := func(amountCell, itemCell, targetCell string) (string, error) {
        amount, err := getIntValue(amountCell)
        if err != nil {
            return "", err // Handle errors appropriately
        }

        if amount != 0 {
            item, _ := getStringValue(itemCell)
            target, _ := getStringValue(targetCell)
            return fmt.Sprintf("You invested %d %s into %s.\n", amount, item, target), nil
        }
        return "", nil // No improvement was made, so no message to add
    }
    
    //Check and format each improvement
    for _, improvement := range []struct {
        amountCell string
        itemCell   string
        targetCell string
    }{
        {fmt.Sprintf("P%d", simHour), fmt.Sprintf("O%d", simHour), fmt.Sprintf("Q%d", simHour)},
        {fmt.Sprintf("S%d", simHour), fmt.Sprintf("R%d", simHour), fmt.Sprintf("T%d", simHour)},
        {fmt.Sprintf("V%d", simHour), fmt.Sprintf("U%d", simHour), fmt.Sprintf("W%d", simHour)},
    } {
        result, err := checkAndFormatImprovement(improvement.amountCell, improvement.itemCell, improvement.targetCell)
        if err != nil {
            return "", err
        }
        actions.WriteString(result)
    }

    return actions.String(), nil
}

package main

import (
    "fmt"
    "strconv"
    "strings"
)

// ... (your other types, constants, ExcelizeInterface, etc.)

func (c *GameLogCmd) stats(outputSheet, outputTextBox string) error {
    var stats strings.Builder

    // Get log hour, defaulting to 72
    logHourStr, err := c.sim.GetCellValue("Overview", "I28")
    if err != nil {
        return fmt.Errorf("error reading log hour: %w", err)
    }
    logHour := 72 // Default value
    if logHourStr != "" {
        logHour, _ = strconv.Atoi(logHourStr)
    }
    hr := logHour + 4 // The hour to read statistics from

    // Helper function to get and format cell values
    getFormattedValue := func(sheet, axis string, format string) (string, error) {
        val, err := c.sim.GetCellValue(sheet, axis)
        if err != nil {
            return "", fmt.Errorf("error reading cell %s!%s: %w", sheet, axis, err)
        }
        if format == "#,##0" || format == "#,###" {
            intVal, err := strconv.Atoi(val)
            if err != nil {
                return "", fmt.Errorf("error converting cell %s!%s value to int: %w", sheet, axis, err)
            }
            // Format the integer with commas
            return fmt.Sprintf(format, intVal), nil
        }
        return fmt.Sprintf(format, val), nil 
    }

    // Function to append a stat line to the builder
    addStatLine := func(label, sheet, axis, format string) error {
        val, err := getFormattedValue(sheet, fmt.Sprintf(axis, hr), format)
        if err != nil {
            return err
        }
        stats.WriteString(fmt.Sprintf("%s:  %s\n", label, val))
        return nil
    }
 // 1. Basic Overview
    stats.WriteString(fmt.Sprintf("The Dominion of Simulated Dominion: Hour %d\nOverview\n", logHour))
    addStatLine("Ruler:", "Overview", "B14", "%s")
    addStatLine("Race:", "Overview", "B14", "%s")
    addStatLine("Land:", "Production", "E%d", "#,###")
    addStatLine("Peasants:", "Population", "C%d", "#,###")
    addStatLine("Draftees:", "Population", "E%d", "#,###")
    addStatLine("Employment:", "Population", "I%d", "%.2f%%") // Multiply by 100 for percentage and format to 2 decimal places
    addStatLine("Networth:", "Production", "G%d", "#,###")

 // 2. Resources
    stats.WriteString("\nResources\n")
    addStatLine("Platinum:", "Production", "H%d", "#,###")
    addStatLine("Food:", "Production", "I%d", "#,##0")
    addStatLine("Lumber:", "Production", "J%d", "#,##0")
    addStatLine("Mana:", "Production", "K%d", "#,##0")
    addStatLine("Ore:", "Production", "L%d", "#,##0")
    addStatLine("Gems:", "Production", "M%d", "#,##0")
    addStatLine("Boats:", "Production", "N%d", "#,##0")

    // 3. Military 
    stats.WriteString("\nMilitary\n")
    stats.WriteString("Morale:  100.00%\n")
    for row := 36; row <= 39; row++ {
        unitLabel, _ := c.sim.GetCellValue("Overview", fmt.Sprintf("A%d", row))
        unitCount, _ := getFormattedValue("Military", fmt.Sprintf("%c%d", 'E'+row-36, hr), "#,##0")
        stats.WriteString(fmt.Sprintf("%s:  %s\n", unitLabel, unitCount))
    }
    addStatLine("Spies:", "Military", "I%d", "#,##0")
    addStatLine("Archspies:", "Military", "J%d", "#,##0")
    addStatLine("Wizards:", "Military", "K%d", "#,##0")
    addStatLine("Archmages:", "Military", "L%d", "#,##0")

    // 4. Additional Stats from Table
    stats.WriteString("\n--------------------------------------------------\n")

    // ... logic to read and format additional stats from the "Log_support" sheet
    
    if _, err := c.sim.SetCellValue(outputSheet, outputTextBox, stats.String()); err != nil {
        return fmt.Errorf("error writing stats to Excel: %w", err)
    }
    
    return nil
}


