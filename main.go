package main

import (
	"encoding/csv"
	"fmt"
	"os"
	"sort"
	"strconv"
	"time"

	"github.com/xuri/excelize/v2"
)

type Record struct {
	StartDate   time.Time
	EndDate     time.Time
	City        string
	NewCases    int
	TotalCases  int
	TotalDeaths int
}

func main() {
	// Open CSV file
	file, err := os.Open("covid_19_indonesia_time_series_all.csv")
	if err != nil {
		fmt.Println("Error:", err)
		return
	}
	defer file.Close()

	// Read CSV file
	reader := csv.NewReader(file)
	records, err := reader.ReadAll()
	if err != nil {
		fmt.Println("Error:", err)
		return
	}

	cityData := make(map[string]*Record)

	// Convert CSV records and group by city
	for _, record := range records[1:] { // skip header
		city := record[2] // Using the 2nd column for the city

		// Exclude "Indonesia" from the list
		if city == "Indonesia" {
			continue
		}

		date, err := time.Parse("1/2/2006", record[0]) // Adjusting to the format "M/D/YYYY"
		if err != nil {
			continue // skip records with invalid date
		}

		newCases, _ := strconv.Atoi(record[4])    // Assuming "New Cases" is the 5th column
		totalCases, _ := strconv.Atoi(record[7])  // Assuming "Total Cases" is the 8th column
		totalDeaths, _ := strconv.Atoi(record[8]) // Assuming "Total Deaths" is the 9th column

		if _, exists := cityData[city]; !exists {
			cityData[city] = &Record{
				StartDate:   date,
				EndDate:     date,
				City:        city,
				NewCases:    newCases,
				TotalCases:  totalCases,
				TotalDeaths: totalDeaths,
			}
		} else {
			cityData[city].NewCases += newCases
			cityData[city].TotalCases += totalCases
			cityData[city].TotalDeaths += totalDeaths
			if date.Before(cityData[city].StartDate) {
				cityData[city].StartDate = date
			}
			if date.After(cityData[city].EndDate) {
				cityData[city].EndDate = date
			}
		}
	}

	// Convert map to slice for sorting
	var data []*Record
	for _, record := range cityData {
		data = append(data, record)
	}

	// Sort data by total cases in descending order
	sort.Slice(data, func(i, j int) bool {
		return data[i].TotalCases > data[j].TotalCases
	})

	// Create a new XLSX file
	f := excelize.NewFile()

	// Write data to XLSX
	f.SetCellValue("Sheet1", "A1", "City")
	f.SetCellValue("Sheet1", "B1", "Total Cases")
	f.SetCellValue("Sheet1", "C1", "New Cases")
	f.SetCellValue("Sheet1", "D1", "Total Deaths")
	for i, record := range data {
		row := i + 2
		f.SetCellValue("Sheet1", fmt.Sprintf("A%d", row), record.City)
		f.SetCellValue("Sheet1", fmt.Sprintf("B%d", row), record.TotalCases)
		f.SetCellValue("Sheet1", fmt.Sprintf("C%d", row), record.NewCases)
		f.SetCellValue("Sheet1", fmt.Sprintf("D%d", row), record.TotalDeaths)
	}

	// Add chart to XLSX
	chart := excelize.Chart{
		Type: excelize.Col3DClustered,
		Series: []excelize.ChartSeries{
			{
				Name:       "Sheet1!$B$1",
				Categories: "Sheet1!$A$2:$A$6",
				Values:     "Sheet1!$B$2:$B$6",
			},
		},
		Title: excelize.ChartTitle{
			Name: "Top 5 Cities with Highest Total Cases",
		},
	}

	if err := f.AddChart("Sheet1", "F5", &chart); err != nil {
		fmt.Println("Error adding chart:", err)
		return
	}

	// Save the XLSX file
	if err := f.SaveAs("output.xlsx"); err != nil {
		fmt.Println("Error:", err)
	}
}
