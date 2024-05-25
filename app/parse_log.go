package main

import (
	"bufio"
	"encoding/json"
	"fmt"
	"os"
	"regexp"
	"runtime/debug"
	"strconv"
	"strings"
)

const (
	BANK         = "bank"
	CONSTRUCTION = "construction"
	DAILY        = "daily"
	DESTRUCTION  = "destruction"
	DRAFTRATE    = "draftrate"
	EXPLORE      = "explore"
	INVEST       = "invest"
	MAGIC        = "magic"
	RELEASE      = "release"
	REZONE       = "rezone"
	TRAIN        = "train"
)

var valuesMap = map[string]string{
	"draftees":       "military_draftees",
	"Draftees":       "military_draftees", // Handle both cases (lower/upper)
	"Spies":          "military_spies",
	"Archspies":      "military_assassins",
	"Wizards":        "military_wizards",
	"Archmages":      "military_archmages",
	"Fire Spirit":    "Fire Sprite",
	"Ice Beast":      "Icebeast",
	"Frost Mage":     "FrostMage",
	"Voodoo Magi":    "Voodoo Mage",
	"Mermen":         "Merman",
	"Sirens":         "Siren",
	"Alchemies":      "alchemy",
	"Barracks":       "barracks",
	"Factories":      "factory",
	"Guilds":         "wizard_guild",
	"Lumber Yards":   "lumberyard",
	"Lumberyards":    "lumberyard",
	"Masonries":      "masonry",
	"Smithies":       "smithy",
	"Ares Call":      "Ares' Call", // Use Go's built-in escape character for single quotes
	"Gaias Blessing": "Gaia's Blessing",
	"Gaias Watch":    "Gaia's Watch",
	"Miners Sight":   "Miner's Sight",
}

type Scanner interface {
	Scan() bool
	Text() string
	Err() error
}

type LogFile interface {
	Close() error
}

type ParseLogFunc func() error

type ActionResultData map[string]int

type ActionResult struct {
	Type string
	Data map[string]int
}

type LogCmd struct {
	currentHour   int
	logPath       string
	scanner       Scanner
	file          LogFile
	currentText   string
	lineNumber    int
	actionResults map[int][]ActionResult
	actions       []ParseLogFunc
}

func NewLogCmd(path string) *LogCmd {
	cmd := &LogCmd{
		logPath:     path,
		currentHour: 0,
		lineNumber:  0,
	}
	cmd.loadFile()
	cmd.initActions()

	return cmd
}

func (c *LogCmd) initActions() {
	c.actions = []ParseLogFunc{
		c.tickAction,
		c.draftrateAction,
		c.releaseUnitAction,
	}
}

func (c *LogCmd) loadFile() {
	file, err := os.Open(c.logPath)
	if err != nil {
		fmt.Printf("Error on reading log file %v\n", err)
		return
	}

	c.file = file
	c.scanner = bufio.NewScanner(file)
}

func (c *LogCmd) Execute() {
	defer c.file.Close()

	c.actionResults = make(map[int][]ActionResult)

	fmt.Println("Parsing...")

	for c.scanner.Scan() {
		if err := c.scanner.Err(); err != nil {
			fmt.Printf("Error scanning file: %v", err)
			return
		}

		c.currentText = strings.TrimSpace(c.scanner.Text())
		c.lineNumber++

		debugLog("Current Line => ", c.currentText)

		if c.currentText == "" {
			continue
		}

		c.executeActions()

		if c.currentHour > 73 {
			break
		}

	}
}

func (c *LogCmd) executeActions() {
	for _, actionFunc := range c.actions {
		err := actionFunc()
		if err != nil {
			fmt.Printf("Error on executing action: %v: CurrentHour: %v Line %v: %v",
				c.currentHour, err, c.lineNumber, c.currentText)

			if cmdVars.debugEnabled {
				debug.PrintStack()
			}
			return
		}

		data, err := json.MarshalIndent(c.actionResults, "", "  ")
		if err != nil {
			fmt.Println("Error marshalling results:", err)
			return
		}
		fmt.Println(string(data))
	}
}

func (c *LogCmd) tickAction() error {
	hourPattern := regexp.MustCompile(`Protection Hour: (\d+)`)
	matches := hourPattern.FindStringSubmatch(c.currentText)

	debugLog("tickAction", matches)

	if len(matches) == 0 {
		return nil
	}

	hour, err := strconv.Atoi(matches[1])
	if err != nil {
		return fmt.Errorf("error parsing hour: %v", err)
	}

	debugLog("TickAction: Parsed Hour", hour)

	if hour <= c.currentHour {
		return fmt.Errorf("hour %d duplicate or out of order", hour)
	}

	c.currentHour = hour - 1

	return nil
}

func (c *LogCmd) addActionResult(result *ActionResult) {
	_, ok := c.actionResults[c.currentHour]

	if !ok {
		c.actionResults[c.currentHour] = []ActionResult{}
	}

	c.actionResults[c.currentHour] = append(c.actionResults[c.currentHour], *result)
}

func (c *LogCmd) draftrateAction() error {
	pattern := regexp.MustCompile(`Draftrate changed to (\d+)%`) // Regexp pattern
	matches := pattern.FindStringSubmatch(c.currentText)

	debugLog("DraftrateAction", pattern, matches)

	if len(matches) == 0 {
		return nil
	}

	rate, err := strconv.Atoi(matches[1])
	if err != nil {
		return fmt.Errorf("error parsing draftrate: %v", err)
	}

	debugLog("Draftrate:", rate)

	result := &ActionResult{
		Type: DRAFTRATE,
		Data: ActionResultData{"value": rate},
	}

	debugLog("Result", result)

	c.addActionResult(result)

	return nil
}

func (c *LogCmd) releaseUnitAction() error {
	pattern := regexp.MustCompile(`You successfully released ([\w\s,]+)`)
	matches := pattern.FindStringSubmatch(c.currentText)

	debugLog("releaseUnitAction", pattern, matches)

	if len(matches) == 0 {
		return nil
	}

	releasedText := matches[1]
	unitPattern := regexp.MustCompile(`(\d+)\s([\w\s]+)`)
	unitMatches := unitPattern.FindAllStringSubmatch(releasedText, -1)

	debugLog("UnitMatches", unitMatches)

	releaseData := make(ActionResultData)

	for _, unitMatch := range unitMatches {
		amount, err := strconv.Atoi(unitMatch[1])
		if err != nil {
			return fmt.Errorf("error parsing released unit amount: %w", err)
		}
		name := strings.TrimSuffix(unitMatch[2], " into the peasantry")

		if mappedName, ok := valuesMap[name]; ok {
			name = mappedName
		}

		var resultKey string
		if strings.HasPrefix(name, "military_") {
			resultKey = strings.TrimPrefix(name, "military_")
		} else {
			resultKey = name
		}

		releaseData[resultKey] = amount

		result := &ActionResult{
			Type: RELEASE,
			Data: releaseData,
		}

		debugLog("Result", result)

		c.addActionResult(result)

	}

	return nil
}
