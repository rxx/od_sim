package main

import (
	"flag"
	"fmt"
	"os"
)

var cmdVars *FlagSetVars

func main() {
	cmdVars = &FlagSetVars{}

	commands := map[string]*flag.FlagSet{
		GenerateLogCmd: cmdVars.GenerateLogCmd(),
		ParseLogCmd:    cmdVars.ParseLogCmd(),
	}

	if len(os.Args) < 2 {
		printUsage(commands)
		os.Exit(1)
	}

	cmdName := os.Args[1]
	cmd, ok := commands[cmdName]
	if !ok {
		fmt.Printf("Unknown command: %s\n", cmdName)
		printUsage(commands)
		os.Exit(1)
	}
	cmd.Parse(os.Args[2:])

	switch cmdName {
	case GenerateLogCmd:
		if cmdVars.simPath == "" {
			cmd.Usage()
			os.Exit(1)
		}

		gameLogCmd := NewGameLog(cmdVars.simPath)
		gameLogCmd.Execute()
	case ParseLogCmd:
		if cmdVars.logPath == "" {
			cmd.Usage()
			os.Exit(1)
		}
	default:
		printUsage(commands)
	}
}

func printUsage(commands map[string]*flag.FlagSet) {
	fmt.Printf("Usage: %s <command> [options]\n", os.Args[0])
	fmt.Println("Available Commands:")

	for _, cmd := range commands {
		cmd.Usage()
	}
}
