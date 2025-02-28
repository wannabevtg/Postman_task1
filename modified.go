package main

import (
	"flag"
	"fmt"
	"log"
	"os"
	"sort"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

// Function to calculate average
func average(rows [][]string, column string) (map[string]float64, error) {

	var sum float64 = 0
	var count int = 0
	col_no, err := excelize.ColumnNameToNumber(column)
	if err != nil {
		log.Fatal(err)
	}

	for i, row := range rows {
		if i == 0 {
			continue
		}
		cell := row[col_no]
		value, err := strconv.ParseFloat(cell, 64)
		if err != nil {
			log.Fatal(err)
		}
		sum += value
		count++
	}
	if count == 0 {
		return nil, fmt.Errorf("no data found in column")
	}
	averages := make(map[string]float64)
	averages[rows[0][col_no]] = sum / float64(count)
	return averages, nil
}

// Function to calculate branchwise average
func branchwise_average(rows [][]string, column string, branch_column string) (map[string]float64, map[string]float64, error) {

	// Initialize branch maps
	br_maps := map[string]int{
		"2024A3PS": 0, "2024A4PS": 0,
		"2024A5PS": 0, "2024A7PS": 0, "2024A8PS": 0,
		"2024ADPS": 0, "2024AAPS": 0,
	}
	br_maps_m := map[string]float64{
		"2024A3PS": 0.0, "2024A4PS": 0.0,
		"2024A5PS": 0.0, "2024A7PS": 0.0, "2024A8PS": 0.0,
		"2024ADPS": 0.0, "2024AAPS": 0.0,
	}
	br_max := map[string]float64{
		"2024A3PS": 0.0, "2024A4PS": 0.0,
		"2024A5PS": 0.0, "2024A7PS": 0.0, "2024A8PS": 0.0,
		"2024ADPS": 0.0, "2024AAPS": 0.0,
	}

	col_no, err := excelize.ColumnNameToNumber(column)
	if err != nil {
		return nil, nil, err
	}
	roll_column, err := excelize.ColumnNameToNumber(branch_column)
	if err != nil {
		return nil, nil, err
	}

	for i, row := range rows {
		if i == 0 {
			continue
		}
		for key := range br_maps {
			if strings.HasPrefix(row[roll_column-1], key) {
				br_maps[key]++
				value, err := strconv.ParseFloat(row[col_no-1], 64)
				if err != nil {
					return nil, nil, err
				}
				br_maps_m[key] += value
				if value > br_max[key] {
					br_max[key] = value
				}
			}
		}
	}

	for key := range br_maps {
		if br_maps[key] > 0 {
			br_maps_m[key] /= (float64)(br_maps[key])
		}
	}

	return br_maps_m, br_max, nil
}

func main() {

	class := flag.Int64("class", 0, "Enter the class number:")
	flag.Parse()
	fmt.Println(*class)
	fmt.Println("Arguments received:", os.Args)

	if len(os.Args) < 2 {
		fmt.Println("filepath missing: C:/Users/vivek/Downloads/CSF111_202425_01_GradeBook_stripped.xlsx")
		return
	}

	// Get the file path from the command-line argument
	filepath := os.Args[2]

	f, err := excelize.OpenFile(filepath)
	if err != nil {
		log.Fatal(err)
	}
	sheetNames := f.GetSheetList()
	// Call branchwise_average
	rows, err := f.GetRows(sheetNames[0])
	if err != nil {
		log.Fatal(err)
	}
	var updated_rows [][]string
	for i, row := range rows {

		if i == 0 {
			updated_rows = append(updated_rows, row)
			continue
		}

		var s float64 = 0.0
		for k := 0; k < 4; k++ {
			value, err := strconv.ParseFloat(row[k+4], 64)
			if err != nil {
				log.Fatal(err)
			}
			s = s + value

		}
		precompre, err_2 := strconv.ParseFloat(row[8], 64)
		if err_2 != nil {
			log.Fatal(err_2)
		}
		compre, err_3 := strconv.ParseFloat(row[9], 64)
		if err_3 != nil {
			log.Fatal(err_3)
		}
		total, err_4 := strconv.ParseFloat(row[10], 64)
		if err_4 != nil {
			log.Fatal(err_4)
		}
		if precompre != s || compre+precompre != total {
			fmt.Println("Totalling error in row ", i)
			continue

		}
		class_check, err := strconv.ParseInt(row[1], 10, 64)

		if err != nil {
			log.Fatal(err)
		}

		if *class == 0 || *class == class_check {
			updated_rows = append(updated_rows, row)
		}
	}
	var store []float64
	var keys []string
	for r := 0; r < 7; r++ {
		connect := make(map[string]float64)
		for j, row := range updated_rows {
			if j == 0 {
				continue
			}
			val_dec, err := strconv.ParseFloat(row[r+4], 64)
			if err != nil {
				log.Fatal(err)
			}
			connect[row[2]] = val_dec
			store = append(store, val_dec)
			keys = append(keys, row[2])
		}
		sort.Slice(keys, func(i, j int) bool {
			return connect[keys[i]] > connect[keys[j]]
		})
		fmt.Println(updated_rows[0][r+4])
		for y := 0; y < 3; y++ {
			fmt.Println("Rank", " ", y+1, " ", keys[y], "  ", connect[keys[y]])
		}
		store = store[:0]
		keys = keys[:0]
		connect = nil
		column_name, err := excelize.ColumnNumberToName(r + 5)
		if err != nil {
			log.Fatal(err)
		}
		brwise_avg_quiz, brwise_max_quiz, err := branchwise_average(updated_rows, column_name, "D")
		if err != nil {
			log.Fatal(err)
		}

		fmt.Println("Branch-wise Average ", updated_rows[0][r+4])
		for key := range brwise_avg_quiz {
			fmt.Println(key, " ", brwise_avg_quiz[key])
		}
		fmt.Println("Branch-wise Max ", updated_rows[0][r+4])
		for key_1 := range brwise_max_quiz {
			fmt.Println(key_1, " ", brwise_max_quiz[key_1])
		}
	}
	for k := 0; k < 7; k++ {

		column_no, err := excelize.ColumnNumberToName(k + 4)
		if err != nil {
			log.Fatal(err)
		}
		average_quiz, err := average(updated_rows, column_no)
		if err != nil {
			log.Fatal(err)
		}
		for avg := range average_quiz {
			fmt.Println(avg, " ", average_quiz[avg])

		}
		// for _, row_1 := range updated_rows {
		// 	fmt.Println(row_1)
		// }

	}
}
