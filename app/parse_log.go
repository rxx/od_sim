package main

import (
	"bufio"
	"fmt"
	"os"
	"strings"
)

type Scanner interface {
	Scan() bool
	Text() string
	Err() error
}

type LogFile interface {
	Close() error
}

type LogCmd struct {
	currentHour int
	logPath     string
	scanner     Scanner
	file        LogFile
	output      strings.Builder
}

func NewLogCmd(path string) *LogCmd {
	cmd := &LogCmd{
		logPath:     path,
		currentHour: 1,
	}
	cmd.loadFile()

	return cmd
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

	for c.scanner.Scan() {
		if err := c.scanner.Err(); err != nil {
			fmt.Println("Error scanning file:", err)
			return
		}

		fmt.Println(c.scanner.Text())
	}
}
