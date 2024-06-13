package main

import (
	"fmt"
	"os"
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
	LastHour     = 73

	// Magic
	GaiasWatch     = "Gaia's Watch"
	MiningStrength = "Mining Strength"
	AresCall       = "Ares' Call"
	MidasTouch     = "Midas Touch"
	Harmony        = "Harmony"

	RacialSpell = "Racial Spell"

	PlatAwardedMult = 4
	LandBonus       = 20
)

var buildingNames = []string{
	"Homes", "Alchemies", "Farms", "Smithies", "Masonries", "Lumber Yards",
	"Ore Mines", "Gryphon Nests", "Factories", "Guard Towers", "Barracks",
	"Shrines", "Towers", "Temples", "Wizard Guilds", "Diamond Mines", "Schools", "Docks",
}

var exploreLands = map[string]string{
	"Plains":    "T",
	"Forest":    "U",
	"Mountains": "V",
	"Hills":     "W",
	"Swamps":    "X",
	"Caverns":   "Y",
	"Water":     "Z",
}

var rezoneLands = map[string]string{
	"Plains":    "L",
	"Forest":    "M",
	"Mountains": "N",
	"Hills":     "O",
	"Swamps":    "P",
	"Caverns":   "Q",
	"Water":     "R",
}

type ActionFunc func() (string, error)

type Sim interface {
	GetCellValue(sheet, cell string, opts ...excelize.Options) (string, error)
	Close() error
}

type GameLogCmd struct {
	currentHour int
	simHour     int
	simPath     string
	resultPath  string
	sim         Sim
	// sim     *excelize.File
	actions []ActionFunc
}

func NewGameLog(path, resultPath string) *GameLogCmd {
	gameLogCmd := &GameLogCmd{
		simPath:    path,
		resultPath: resultPath,
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
		c.castMagicSpells,
		c.unlockTechAction,
		c.dailtyPlatinumAction,
		c.tradeResources,
		c.exploreAction,
		c.dailyLandAction,
		c.destroyBuildingsAction,
		c.rezoneAction,
		c.constructionAction,
		c.trainUnitsAction,
		c.improvementsAction,
	}
}

func (c *GameLogCmd) validateSim() error {
	//	c.draftRateAction // 90 max, 0 min Military#Z4 to Z76

	//	c.releaseUnitsAction // release not more than units exists
	// Military E, F, G, H, I, J, K, L >= 0
	//	c.castMagicSpells // cast not more than mana available
	// Production#K15 >= 0
	//	c.dailtyPlatinumAction
	// only once in 24 hour range
	// 4-27, 25-48, 49-23
	//	c.tradeResources // all traded resources are valid
	// - 1 lumber = 0.5 platinum
	//- 1 lumber = 0.5 ore
	//- 1 platinum = 0.5 lumber
	//- 1 platinum = 0.5 ore

	//- 1 gem = 1 food
	//- 1 gem = 2 ore
	//- 1 gem = 2 lum
	//- 1 gem = 2 platinum
	//	c.exploreAction // not more than plat, draft available
	// Pipulation L4-76 >= 0

	//	c.dailyLandAction // only once in // 4-27, 25-48, 49-23
	// c.destroyBuildingsAction // no more than available
	// Construction >= 0
	// AY to BQ
	//	c.rezoneAction,
	// c.constructionAction // land available
	// Construction G to M >= 0
	// Production H to M >= 0
	//	c.trainUnitsAction //better left more than 400 platinum after training
	//	c.improvementsAction,
	return nil
}

func (c *GameLogCmd) readConst(cell string) (float64, error) {
	value, err := c.readFloatValue(Constants, cell, "error reading const")
	if err != nil {
		return 0, err
	}
	return value, nil
}

func (c *GameLogCmd) wrapHour(cellCol string) string {
	return c.wrapHourAs(cellCol, c.simHour)
}

func (c *GameLogCmd) wrapHourAs(cellCol string, hour int) string {
	return fmt.Sprintf("%s%d", cellCol, hour)
}

func (c *GameLogCmd) readLandSize() (int, error) {
	value, err := c.readIntValue(Explore, c.wrapHour("B"), "error reading land size")
	if err != nil {
		return 0, err
	}
	return value, nil
}

// Starting at row 4 because of extra added row (due to uniform table headers)
func (c *GameLogCmd) setCurrentHour(hr int) {
	c.currentHour = hr - 1
	c.simHour = hr + 3
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
		return "", WrapError(err, errorMsg)
	}

	return strings.TrimSpace(value), nil
}

func (c *GameLogCmd) readIntValue(sheet, cell, errorMsg string) (int, error) {
	value, err := c.readValue(sheet, cell, errorMsg)
	if err != nil {
		return 0, err
	}

	if value == "" {
		return 0, nil
	}

	// Remove commas (thousands separators) from the string
	value = strings.Trim(value, "→ ")
	value = strings.ReplaceAll(value, ",", "")
	digit, err := strconv.Atoi(value)
	if err != nil {
		return 0, WrapError(err, errorMsg)
	}
	return digit, nil
}

func (c *GameLogCmd) readFloatValue(sheet, cell, errorMsg string) (float64, error) {
	value, err := c.readValue(sheet, cell, errorMsg)
	if err != nil {
		return 0, err
	}

	if value == "" {
		return 0, nil
	}

	digit, err := strconv.ParseFloat(value, 64)
	if err != nil {
		return 0, WrapError(err, errorMsg)
	}
	return digit, nil
}

func (c *GameLogCmd) Execute() {
	defer c.sim.Close()
	var sb strings.Builder

	if cmdVars.hour > 0 {
		c.setCurrentHour(cmdVars.hour)
		result, err := c.executeActions()
		if err != nil {
			fmt.Println(err)
		}
		if result != "" {
			sb.WriteString(result)
		}
	} else {
		for hr := 1; hr <= LastHour; hr++ {
			c.setCurrentHour(hr)
			result, err := c.executeActions()
			if err != nil {
				fmt.Println(err)
				break
			}
			if result == "" {
				continue
			}

			sb.WriteString(result)
			sb.WriteString("\n")
		}
	}

	c.PrintResult(sb.String())
}

func (c *GameLogCmd) PrintResult(result string) {
	if c.resultPath == "" || c.resultPath == "std" {
		fmt.Println(result)
		return
	}

	err := os.WriteFile(c.resultPath, []byte(result), 0644)
	if err != nil {
		fmt.Println("Error writing to file:", err)
		return
	}

	fmt.Printf("Successfully wrote result to %s\n", c.resultPath)
}

func (c *GameLogCmd) executeActions() (string, error) {
	var sb strings.Builder

	for _, actionFunc := range c.actions {
		result, err := actionFunc()
		if err != nil {
			return "", fmt.Errorf("error on executing actions: %v", err)
		}
		if result == "" {
			continue
		}

		sb.WriteString(result)

		if !strings.HasSuffix(result, "\n") {
			sb.WriteString("\n")
		}
	}

	stringResult := sb.String()

	// Skip if we have only timeline without actions
	if strings.HasSuffix(stringResult, "==\n") {
		stringResult = ""
		sb.Reset()
	}

	return stringResult, nil
}

func (c *GameLogCmd) tickAction() (string, error) {
	const dateCell = "B15"
	localTimeCell := c.wrapHour("BY")
	domTimeCell := c.wrapHour("BZ")

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
				return "", WrapError(err, "error parsing date: %w")
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
	timeline.WriteString(" ) ======\n")

	return timeline.String(), nil
}

func (c *GameLogCmd) draftRateAction() (string, error) {
	currentRateCell := c.wrapHour("Y")
	previousRateCell := c.wrapHourAs("Z", c.simHour-1)

	currentRateStr, err := c.readValue(Military, currentRateCell, "error reading current draftrate")
	if err != nil {
		return "", err
	}

	previousRateStr, err := c.readValue(Military, previousRateCell, "error reading previous draftrate")
	if err != nil {
		return "", err
	}

	var buf strings.Builder

	if currentRateStr == "" || currentRateStr == previousRateStr {
		return "", nil
	}

	buf.WriteString("Draftrate changed to ")
	buf.WriteString(currentRateStr)
	buf.WriteString(".\n")

	return buf.String(), nil
}

func (c *GameLogCmd) releaseUnitsAction() (string, error) {
	// Read unit names and unit counts
	cols := []string{"AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE"}

	var sb strings.Builder
	sb.WriteString("You successfully released ")

	addedItems := 0
	for _, col := range cols {
		name, err := c.readValue(Military, c.wrapHourAs(col, 2), "error reading unit name")
		if err != nil {
			return "", err
		}

		value, err := c.readIntValue(Military, c.wrapHour(col), "error reading unit value")
		if err != nil {
			return "", err
		}

		if value == 0 {
			continue
		}

		if addedItems > 0 {
			sb.WriteString(", ")
		}

		sb.WriteString(fmt.Sprintf("%d %s", value, name))
		addedItems++
	}

	if addedItems == 0 {
		sb.Reset()
	} else {
		sb.WriteString(".\n")
	}

	// Read draftees count from AW column
	drafteesCell := c.wrapHour("AW")
	draftees, err := c.readIntValue(Military, drafteesCell, "error reading draftees value")
	if err != nil {
		return "", err
	}

	if draftees > 0 {
		sb.WriteString(fmt.Sprintf("You successfully released %d draftees into the peasantry.\n", draftees))
	}

	return sb.String(), nil
}

func (c *GameLogCmd) castMagicSpells() (string, error) {
	var sb strings.Builder

	landBonusVal, err := c.readIntValue(Explore, c.wrapHour("S"), "error on reading explore cell")
	if err != nil {
		return "", err
	}

	landSize, err := c.readLandSize()
	if err != nil {
		return "", err
	}

	checkAndCastSpell := func(spellName, magicCol, multCell string, isRacial bool) error {
		if isRacial {
			spellName = RacialSpell
		}
		magicCell := c.wrapHour(magicCol)
		magicVal, err := c.readIntValue(Magic, magicCell, "error on reading magic cell")
		if err != nil {
			return err
		}

		multVal, err := c.readConst(multCell)
		if err != nil {
			multVal = 2
			// return err
		}

		mana := 0
		// if land bonus received
		if landBonusVal != 0 && magicVal != 0 {
			mana = FloatToInt((float64(landSize) - LandBonus) * multVal)
		} else if magicVal != 0 {
			mana = FloatToInt(float64(landSize) * multVal)
		} else {
			return nil // No spell was cast, so no message to add
		}

		sb.WriteString(fmt.Sprintf("Your wizards successfully cast %s at a cost of %d mana.\n", spellName, mana))

		return nil
	}

	spells := []struct {
		name     string
		cell     string
		mult     string
		isRacial bool
	}{
		{GaiasWatch, "G", "B75", false},
		{MiningStrength, "H", "B76", false},
		{AresCall, "I", "B77", false},
		{MidasTouch, "J", "B78", false},
		{Harmony, "K", "B79", false},
		{"", "L", "B80", true},
		{"", "M", "B80", true},
		{"", "N", "B80", true},
		{"", "O", "B80", true},
		{"", "P", "B80", true},
		{"", "Q", "B80", true},
		{"", "R", "B80", true},
		{"", "S", "B80", true},
		{"", "T", "B80", true},
		{"", "U", "B80", true},
	}

	// Check and cast each spell
	for _, spell := range spells {
		if err := checkAndCastSpell(spell.name, spell.cell, spell.mult, spell.isRacial); err != nil {
			return "", WrapError(err, "error on casting magic spell")
		}
	}
	return sb.String(), nil
}

func (c *GameLogCmd) unlockTechAction() (string, error) {
	// Check if a tech was unlocked
	techUnlocked, err := c.readIntValue(Techs, c.wrapHour("K"), "error reading tech status")
	if err != nil {
		return "", err
	}

	if techUnlocked > 0 {
		techName, err := c.readValue(Techs, c.wrapHour("CA"), "error reading tech name")
		if err != nil {
			return "", err
		}

		return fmt.Sprintf("You have unlocked %s.\n", techName), nil
	}

	return "", nil
}

func (c *GameLogCmd) dailtyPlatinumAction() (string, error) {
	platChecked, err := c.readIntValue(Production, c.wrapHour("C"), "error reading platinum bonus")
	if err != nil {
		return "", err
	}
	if platChecked == 0 {
		return "", nil
	}

	populationValue, err := c.readIntValue(Population, c.wrapHour("C"), "error reading population")
	if err != nil {
		return "", err
	}

	platinumAwarded := populationValue * PlatAwardedMult
	return fmt.Sprintf("You have been awarded with %d platinum.\n", platinumAwarded), nil
}

func (c *GameLogCmd) tradeResources() (string, error) {
	var sb strings.Builder

	plat, err := c.readIntValue(Production, c.wrapHour("BC"), "can't read platinum value for trading")
	if err != nil {
		return "", err
	}

	lumber, err := c.readIntValue(Production, c.wrapHour("BD"), "can't read lumber value for trading")
	if err != nil {
		return "", err
	}

	ore, err := c.readIntValue(Production, c.wrapHour("BE"), "can't read ore value for trading")
	if err != nil {
		return "", err
	}

	gems, err := c.readIntValue(Production, c.wrapHour("BF"), "can't read gems value for trading")
	if err != nil {
		return "", err
	}

	if plat == 0 && lumber == 0 && ore == 0 && gems == 0 { // Check if any exchange happened
		return "", nil
	}
	var tradedItems []string
	var receivedItems []string

	addItem := func(item string, amount int) {
		formatValue := func(value int) string {
			return fmt.Sprintf("%d %s", value, item)
		}

		if amount < 0 {
			tradedItems = append(tradedItems, formatValue(-amount))
		} else if amount > 0 {
			receivedItems = append(receivedItems, formatValue(amount))
		}
	}

	addItem("platinum", plat)
	addItem("lumber", lumber)
	addItem("ore", ore)
	addItem("gems", gems)

	// Construct the action message
	if len(tradedItems) > 0 {
		sb.WriteString(strings.Join(tradedItems, " and ") + " have been traded for ")
	}
	if len(receivedItems) > 0 {
		sb.WriteString(strings.Join(receivedItems, " and ") + ".\n")
	}

	return sb.String(), nil
}

func (c *GameLogCmd) exploreAction() (string, error) {
	var sb strings.Builder

	sb.WriteString("Exploration for ")

	addedItems := 0
	// Read exploration counts for each land type
	for landType, col := range exploreLands {
		cell := c.wrapHour(col)
		value, err := c.readIntValue(Explore, cell, "error on reading land amount")
		if err != nil {
			return "", err
		}

		if value == 0 {
			continue
		}

		if addedItems > 0 {
			sb.WriteString(", ")
		}

		sb.WriteString(fmt.Sprintf("%d %s", value, landType))
		addedItems++
	}

	if addedItems == 0 {
		sb.Reset()
		return "", nil
	}

	// Read cost values
	platCost, err := c.readIntValue(Explore, c.wrapHour("AH"), "error reading explore plat cost")
	if err != nil {
		return "", nil
	}
	drafteeCost, err := c.readIntValue(Explore, c.wrapHour("AI"), "error reading explore draftees costs")
	if err != nil {
		return "", nil
	}

	sb.WriteString(fmt.Sprintf(" begun at a cost of %d platinum and %d draftees.\n", platCost, drafteeCost))

	return sb.String(), nil
}

func (c *GameLogCmd) dailyLandAction() (string, error) {
	landBonus, err := c.readIntValue(Explore, c.wrapHour("S"), "error on reading land bonus value")
	if err != nil {
		return "", err
	}

	if landBonus == 0 {
		return "", nil
	}

	landType, err := c.readValue(Overview, "B70", "error reading land type")
	if err != nil {
		return "", err
	}

	return fmt.Sprintf("You have been awarded with %d %s.\n", LandBonus, landType), nil
}

func (c *GameLogCmd) destroyBuildingsAction() (string, error) {
	var sb strings.Builder

	cols := []string{
		"BW", "BX", "BY", "BZ", "CA", "CB", "CD", "CE", "CF", "CG",
		"CH", "CI", "CJ", "CK", "CL", "CM", "CN", "CO",
	}

	sb.WriteString("Destruction of ")
	addedItems := 0

	for index, col := range cols {
		name := buildingNames[index]
		value, err := c.readIntValue(Construction, c.wrapHour(col), "error on reading destroy value")
		if err != nil {
			return "", err
		}

		if value == 0 {
			continue
		}

		if addedItems > 0 {
			sb.WriteString(", ")
		}

		sb.WriteString(fmt.Sprintf("%d %s", value, name))
		addedItems++
	}

	if addedItems == 0 {
		sb.Reset()
		return "", nil
	}

	sb.WriteString(" is complete.\n")

	return sb.String(), nil
}

func (c *GameLogCmd) rezoneAction() (string, error) {
	var sb strings.Builder

	platCost, err := c.readIntValue(Rezone, c.wrapHour("Y"), "error on reading rezone cost")
	if err != nil {
		return "", err
	}
	if platCost == 0 {
		return "", nil
	}

	sb.WriteString(fmt.Sprintf("Rezoning begun at a cost of %d platinum. The changes in land are as following: ", platCost))

	addedItems := 0
	for landType, col := range rezoneLands {
		value, err := c.readIntValue(Rezone, c.wrapHour(col), "error on reading rezone value")
		if err != nil {
			return "", err
		}

		if value == 0 {
			continue
		}
		if addedItems > 0 {
			sb.WriteString(", ")
		}

		sb.WriteString(fmt.Sprintf("%d %s", value, landType))
		addedItems++
	}

	sb.WriteString(".\n")

	return sb.String(), nil
}

func (c *GameLogCmd) constructionAction() (string, error) {
	var sb strings.Builder
	sb.WriteString("Construction of ")

	cols := []string{"O", "P", "Q", "R", "S", "T", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG"}

	addedItems := 0
	for index, col := range cols {
		name := buildingNames[index]
		value, err := c.readIntValue(Construction, c.wrapHour(col), "error on reading construction value")
		if err != nil {
			return "", err
		}
		if value == 0 {
			continue
		}
		if addedItems > 0 {
			sb.WriteString(", ")
		}

		sb.WriteString(fmt.Sprintf("%d %s", value, name))
		addedItems++
	}

	if addedItems == 0 {
		sb.Reset()
		return "", nil
	}

	// Read cost values
	platCost, err := c.readIntValue(Construction, c.wrapHour("AQ"), "error reading platinum cost")
	if err != nil {
		return "", err
	}

	lumberCost, err := c.readIntValue(Construction, c.wrapHour("AR"), "error reading lumber cost")
	if err != nil {
		return "", err
	}

	sb.WriteString(fmt.Sprintf(" started at a cost of %d platinum and %d lumber.\n", platCost, lumberCost))

	return sb.String(), nil
}

func (c *GameLogCmd) trainUnitsAction() (string, error) {
	var sb strings.Builder
	sb.WriteString("Training of ")
	cols := []string{"AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN"}
	addedItems := 0

	drafteesCount := 0
	spiesCount := 0
	wizardCount := 0
	for index, col := range cols {
		name, err := c.readValue(Military, c.wrapHourAs(col, 2), "error reading unit name cell")
		if err != nil {
			return "", err
		}

		value, err := c.readIntValue(Military, c.wrapHour(col), "error reading unit value cell")
		if err != nil {
			return "", err
		}

		if value == 0 {
			continue
		}

		switch index {
		case 5:
			spiesCount += value
		case 7:
			wizardCount += value
		default:
			drafteesCount += value
		}

		if addedItems > 0 {
			sb.WriteString(", ")
		}

		sb.WriteString(fmt.Sprintf("%d %s", value, name))
		addedItems++
	}

	if addedItems == 0 {
		return "", nil
	}

	platCost, err := c.readIntValue(Military, c.wrapHour("AR"), "error reading platinum training cost")
	if err != nil {
		return "", err
	}
	oreCost, err := c.readIntValue(Military, c.wrapHour("AS"), "error reading ore training cost")
	if err != nil {
		return "", err
	}

	sb.WriteString(fmt.Sprintf(" begun at a cost of %d platinum, %d ore, %d draftees, %d spies, and %d wizards.\n",
		platCost, oreCost, drafteesCount, spiesCount, wizardCount))

	return sb.String(), nil
}

func (c *GameLogCmd) improvementsAction() (string, error) {
	var sb strings.Builder

	checkAndFormatImprovement := func(amountCol, resourceCol, targetCol string) (string, error) {
		amount, err := c.readIntValue(Imps, c.wrapHour(amountCol), "error on read amout cell")
		if err != nil {
			return "", err
		}
		if amount == 0 {
			return "", nil
		}

		resource, err := c.readValue(Imps, c.wrapHour(resourceCol), "error on read resource cell")
		if err != nil {
			return "", err
		}

		target, err := c.readValue(Imps, c.wrapHour(targetCol), "error on read improvement cell")
		if err != nil {
			return "", err
		}

		return fmt.Sprintf("You invested %d %s into %s.\n", amount, resource, target), nil
	}

	improvments := []struct {
		amountCell   string
		resourceCell string
		targetCell   string
	}{
		{"P", "O", "Q"},
		{"S", "R", "T"},
		{"V", "U", "W"},
	}

	for _, imp := range improvments {
		result, err := checkAndFormatImprovement(imp.amountCell, imp.resourceCell, imp.targetCell)
		if err != nil {
			return "", err
		}

		sb.WriteString(result)
	}

	return sb.String(), nil
}

//
// // ... (your other types, constants, ExcelizeInterface, etc.)
//
// func (c *GameLogCmd) stats(outputSheet, outputTextBox string) error {
// 	var stats strings.Builder
//
// 	// Get log hour, defaulting to 72
// 	logHourStr, err := c.sim.GetCellValue("Overview", "I28")
// 	if err != nil {
// 		return fmt.Errorf("error reading log hour: %w", err)
// 	}
// 	logHour := 72 // Default value
// 	if logHourStr != "" {
// 		logHour, _ = strconv.Atoi(logHourStr)
// 	}
// 	hr := logHour + 4 // The hour to read statistics from
//
// 	// Helper function to get and format cell values
// 	getFormattedValue := func(sheet, axis string, format string) (string, error) {
// 		val, err := c.sim.GetCellValue(sheet, axis)
// 		if err != nil {
// 			return "", fmt.Errorf("error reading cell %s!%s: %w", sheet, axis, err)
// 		}
// 		if format == "#,##0" || format == "#,###" {
// 			intVal, err := strconv.Atoi(val)
// 			if err != nil {
// 				return "", fmt.Errorf("error converting cell %s!%s value to int: %w", sheet, axis, err)
// 			}
// 			// Format the integer with commas
// 			return fmt.Sprintf(format, intVal), nil
// 		}
// 		return fmt.Sprintf(format, val), nil
// 	}
//
// 	// Function to append a stat line to the builder
// 	addStatLine := func(label, sheet, axis, format string) error {
// 		val, err := getFormattedValue(sheet, fmt.Sprintf(axis, hr), format)
// 		if err != nil {
// 			return err
// 		}
// 		stats.WriteString(fmt.Sprintf("%s:  %s\n", label, val))
// 		return nil
// 	}
// 	// 1. Basic Overview
// 	stats.WriteString(fmt.Sprintf("The Dominion of Simulated Dominion: Hour %d\nOverview\n", logHour))
// 	addStatLine("Ruler:", "Overview", "B14", "%s")
// 	addStatLine("Race:", "Overview", "B14", "%s")
// 	addStatLine("Land:", "Production", "E%d", "#,###")
// 	addStatLine("Peasants:", "Population", "C%d", "#,###")
// 	addStatLine("Draftees:", "Population", "E%d", "#,###")
// 	addStatLine("Employment:", "Population", "I%d", "%.2f%%") // Multiply by 100 for percentage and format to 2 decimal places
// 	addStatLine("Networth:", "Production", "G%d", "#,###")
//
// 	// 2. Resources
// 	stats.WriteString("\nResources\n")
// 	addStatLine("Platinum:", "Production", "H%d", "#,###")
// 	addStatLine("Food:", "Production", "I%d", "#,##0")
// 	addStatLine("Lumber:", "Production", "J%d", "#,##0")
// 	addStatLine("Mana:", "Production", "K%d", "#,##0")
// 	addStatLine("Ore:", "Production", "L%d", "#,##0")
// 	addStatLine("Gems:", "Production", "M%d", "#,##0")
// 	addStatLine("Boats:", "Production", "N%d", "#,##0")
//
// 	// 3. Military
// 	stats.WriteString("\nMilitary\n")
// 	stats.WriteString("Morale:  100.00%\n")
// 	for row := 36; row <= 39; row++ {
// 		unitLabel, _ := c.sim.GetCellValue("Overview", fmt.Sprintf("A%d", row))
// 		unitCount, _ := getFormattedValue("Military", fmt.Sprintf("%c%d", 'E'+row-36, hr), "#,##0")
// 		stats.WriteString(fmt.Sprintf("%s:  %s\n", unitLabel, unitCount))
// 	}
// 	addStatLine("Spies:", "Military", "I%d", "#,##0")
// 	addStatLine("Archspies:", "Military", "J%d", "#,##0")
// 	addStatLine("Wizards:", "Military", "K%d", "#,##0")
// 	addStatLine("Archmages:", "Military", "L%d", "#,##0")
//
// 	// 4. Additional Stats from Table
// 	stats.WriteString("\n--------------------------------------------------\n")
//
// 	// ... logic to read and format additional stats from the "Log_support" sheet
//
// 	if _, err := c.sim.SetCellValue(outputSheet, outputTextBox, stats.String()); err != nil {
// 		return fmt.Errorf("error writing stats to Excel: %w", err)
// 	}
//
// 	return nil
// }
