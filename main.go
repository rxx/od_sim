  package main

  import (
      "fmt"

      "github.com/xuri/excelize/v2"
  )

  func main() {
      f, err := excelize.OpenFile("data/sim.xlsm")
      if err != nil {
          fmt.Println(err)
          return
      }
      defer func() {
          // Close the spreadsheet.
          if err := f.Close(); err != nil {
              fmt.Println(err)
          }
      }()
      // Get value from cell by given worksheet name and cell reference.
      cell, err := f.GetCellValue("Overview", "B14")
      if err != nil {
          fmt.Println(err)
          return
      }
      fmt.Println(cell)
    
  }