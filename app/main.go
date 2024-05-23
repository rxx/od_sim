package main

import (
	"flag"
	"fmt"
	"os"
)

func main() {
	generateLogCmd := flag.NewFlagSet("generate_log", flag.ExitOnError)
	sheetFilePath := generateLogCmd.String("sim", "", "Path to the sim file")

	if len(os.Args) < 2 {
		fmt.Println("Expected a command: generate_log")
		os.Exit(1)
	}

	switch os.Args[1] {
	case "generate_log":
		generateLogCmd.Parse(os.Args[2:])
		if *sheetFilePath == "" {
			generateLogCmdUsage(generateLogCmd)
			os.Exit(1)
		}

		executeGenerateLogCmd(*sheetFilePath)
	default:
		fmt.Println("Unknown command")
		return
	}
}

func generateLogCmdUsage(cmd *flag.FlagSet) {
	fmt.Printf("Usage of %s generate_log:\n", os.Args[0])
	cmd.PrintDefaults() // This will print all defined flags and their descriptions
	fmt.Println("\nExample:")
	fmt.Printf("  %s generate_log -sim data.xlsx\n", os.Args[0])
}
