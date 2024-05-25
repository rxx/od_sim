package main

import (
	"flag"
	"fmt"
	"os"
)

type FlagSetVars struct {
	debugEnabled bool
	simPath      string
	logPath      string
}

const (
	GenerateLogCmd = "generate_log"
)

func (c *FlagSetVars) GenerateLogCmd() *flag.FlagSet {
	cmd := flag.NewFlagSet(GenerateLogCmd, flag.ExitOnError)
	cmd.BoolVar(&c.debugEnabled, "debug", false, "Enable debug logging")
	cmd.StringVar(&c.simPath, "sim", "", "Path to the sim file")
	cmd.Usage = func() {
		fmt.Printf("Usage of %s generate_log:\n", os.Args[0])
		cmd.PrintDefaults() // This will print all defined flags and their descriptions
		fmt.Println("\nExample:")
		fmt.Printf("  %s generate_log -sim data.xlsm\n", os.Args[0])
	}

	return cmd
}
