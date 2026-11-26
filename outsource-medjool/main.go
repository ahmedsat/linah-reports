package main

//TODO - dont remove old output file unless sucsses

import (
	"bytes"
	_ "embed"
	"errors"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

var out *excelize.File

const outputFileName = "output.xlsx"

var (
	dateRE = regexp.MustCompile(`(\d{2})[-_](\d{2})(?:[-_](\d{2,4}))?`)
	osRE   = regexp.MustCompile(`(?i)os\s*(\d{1,2})\b`)
)

type supplier struct {
	Name string
	City string
}

var supplierCodes map[string]supplier

//go:embed output-template.xlsx
var templateXLSX []byte

func main() {
	if err := run(); err != nil {
		log.Print(err)
		appendErrorLog(err)
		os.Exit(1)
	}
}

func appendErrorLog(err error) {
	f, openErr := os.OpenFile("error.txt", os.O_APPEND|os.O_CREATE|os.O_WRONLY, 0o644)
	if openErr != nil {
		return
	}
	defer f.Close()

	fmt.Fprintf(f, "%s %v\n", time.Now().Format("2006-01-02 15:04:05"), err)
}

func run() error {
	var err error

	out, err = excelize.OpenReader(bytes.NewReader(templateXLSX))
	if err != nil {
		return err
	}
	defer func() {
		if out != nil {
			out.Close()
		}
	}()

	if err := loadSuppliers(); err != nil {
		return err
	}

	entries, err := os.ReadDir("./")
	if err != nil {
		return err
	}

	for _, e := range entries {
		name := e.Name()
		if strings.Contains(name, "Zone.Identifier") { //NOTE - ignore the shitty files Microslop creates
			continue
		}

		if strings.HasPrefix(name, "اوزان لينه للتنميه") {
			if err := processFile(name); err != nil {
				return err
			}
		}
	}

	if lastRawIndex == 2 && lastGradeIndex == 2 {
		return fmt.Errorf("no data rows were written; %s was not updated", outputFileName)
	}

	// Excelize does not update the worksheet <dimension> when rows are duplicated,
	// so SheetJS (used by the website) reads only the template's first data row.
	if err := out.SetSheetDimension("Raw os", fmt.Sprintf("A1:H%d", lastRawIndex-1)); err != nil {
		return err
	}
	if err := out.SetSheetDimension("Grades os ", fmt.Sprintf("A1:F%d", lastGradeIndex-1)); err != nil {
		return err
	}

	tmpName := "output.tmp.xlsx"
	if err := out.SaveAs(tmpName); err != nil {
		return err
	}
	if err := out.Close(); err != nil {
		return err
	}
	out = nil

	return replaceOutput(tmpName)
}

func replaceOutput(tmpName string) error {
	if err := os.Rename(tmpName, outputFileName); err == nil {
		return nil
	}

	// Windows does not overwrite an existing file with os.Rename.
	if err := os.Remove(outputFileName); err != nil && !errors.Is(err, os.ErrNotExist) {
		return err
	}
	return os.Rename(tmpName, outputFileName)
}

func processFile(path string) error {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return fmt.Errorf("can not open file %s: %w", path, err)
	}
	defer f.Close()

	for _, sheet := range f.GetSheetList() {
		if code := ExtractOSCode(sheet); code != "" {
			if err := handelSheet(f, sheet, code); err != nil {
				return err
			}
		}
	}
	return nil
}

func loadSuppliers() error {
	rows, err := out.GetRows("اكواد الموردين")
	if err != nil {
		return err
	}

	supplierCodes = make(map[string]supplier, len(rows))
	for _, row := range rows[1:] {
		if len(row) < 3 {
			continue
		}
		code := strings.TrimSpace(row[0])
		if code == "" {
			continue
		}
		supplierCodes[code] = supplier{
			Name: strings.TrimSpace(row[1]),
			City: strings.TrimSpace(row[2]),
		}
	}
	return nil
}

func setSupplierCells(sheet, code string, row int) error {
	info := supplierCodes[code]
	if err := out.SetCellStr(sheet, fmt.Sprintf("E%d", row), info.Name); err != nil {
		return err
	}
	if err := out.SetCellStr(sheet, fmt.Sprintf("F%d", row), info.City); err != nil {
		return err
	}
	return nil
}

func ExtractDateFromFilename(filename string) (string, error) {
	match := dateRE.FindStringSubmatch(filepath.Base(filename))
	if match == nil {
		return "", fmt.Errorf("no date found in filename: %s", filename)
	}

	day, month, year := match[1], match[2], match[3]
	var (
		t   time.Time
		err error
	)

	switch len(year) {
	case 0:
		t, err = time.Parse("02-01", day+"-"+month)
		if err == nil {
			t = t.AddDate(time.Now().Year()-t.Year(), 0, 0)
		}
	case 2:
		t, err = time.Parse("02-01-06", day+"-"+month+"-"+year)
	case 4:
		t, err = time.Parse("02-01-2006", day+"-"+month+"-"+year)
	default:
		return "", fmt.Errorf("unsupported year %q in filename: %s", year, filename)
	}
	if err != nil {
		return "", err
	}

	return t.Format("02/01/2006"), nil
}

type grade int

const (
	unset grade = iota
	grade1
	grade2
	grade3
	grade4
	grade4Dust
	goldenCrown
	dry
	rutab
	underProcessing
	productionWaste
	waste
	compostWaste
	hashaf

	gradesCount
)

func (g grade) String() string {
	switch g {
	case grade1:
		return "Grade 1"
	case grade2:
		return "Grade 2"
	case grade3:
		return "Grade 3"
	case grade4:
		return "Grade 4"
	case grade4Dust:
		return "Grade 4 (Dust Remains)"
	case goldenCrown:
		return "Golden Crown"
	case rutab:
		return "Rutab"
	case underProcessing:
		return "Under Processing"
	case waste:
		return "Waste"
	case dry:
		return "Dry"
	case productionWaste:
		return "Production Waste"
	case compostWaste:
		return "Compost Waste"
	case hashaf:
		return "Hashaf"

	default:
		return fmt.Sprintf("Unknown Grade: %d", g)
	}
}

func ExtractOSCode(s string) string {
	match := osRE.FindStringSubmatch(s)

	if match == nil {
		return ""
	}

	return "os" + strings.ToLower(match[1])
}

func handelSheet(file *excelize.File, sheet, code string) error {
	const (
		rowWidth = 13
		startRow = 4
	)

	date, err := ExtractDateFromFilename(file.Path)
	if err != nil {
		return err
	}
	if _, ok := supplierCodes[code]; !ok {
		return fmt.Errorf("unknown supplier code %q for date %s", code, date)
	}
	rawQuantity := 0.0
	grades := [gradesCount]float64{}

	rows, err := file.GetRows(sheet)
	if err != nil {
		return err
	}

	for _, row := range rows[startRow:] {
		for len(row) < rowWidth {
			row = append(row, "")
		}
		if v, err := strconv.ParseFloat(strings.TrimSpace(row[0]), 64); err == nil && v == 0 {
			break
		}

		row[1] = strings.TrimSpace(row[1])
		if row[1] == "" {
			row[1] = code
		}

		row[7] = strings.TrimSpace(row[7])
		if row[7] == "" {
			row[7] = code
		}

		if row[1] != "" {
			if code != row[1] {
				return fmt.Errorf("conflict codes `%s` and `%s`", code, row[1])
			}
			q, err := strconv.ParseFloat(row[3], 64)
			if err != nil {
				return fmt.Errorf("quantity \"%s\", is not a number", row[3])
			}
			rawQuantity += q
		}

		if row[7] != "" {
			if code != row[7] {
				return fmt.Errorf("conflict codes `%s` and `%s`", code, row[7])
			}
			gq, err := strconv.ParseFloat(row[9], 64)
			if err != nil {
				return fmt.Errorf("quantity \"%s\", is not a number", row[9])
			}

			gradeText := strings.Join(strings.Fields(row[12]), " ")
			if gradeText == "" {
				if gq != 0 {
					return fmt.Errorf("missing grade for date %s code %s quantity %q", date, code, row[9])
				}
				continue
			}

			grade := unset
			switch gradeText {
			case "مطوق":
				grade = goldenCrown
			case "1":
				grade = grade1
			case "2":
				grade = grade2
			case "3":
				grade = grade3
			case "4":
				grade = grade4
			case "4 بقايا اتربه", "بقايا اتربه":
				grade = grade4Dust
			case "مرطب":
				grade = rutab
			case "هالك":
				grade = waste
			case "جاف":
				grade = dry
			case "تحت التصنيع":
				grade = underProcessing
			case "هالك تصنيع", "Production Waste", "production waste", "PRODUCTION WASTE":
				grade = productionWaste
			case "هالك كمبوست", "Compost Waste", "compost waste", "COMPOST WASTE":
				grade = compostWaste
			case "حشف", "Hashaf", "hashaf", "HASHAF":
				grade = hashaf
			default:
				return fmt.Errorf("grade `%s` is unknown", gradeText)
			}

			grades[grade] += gq
		}
	}

	if err := AddRaw(date, code, rawQuantity); err != nil {
		return err
	}

	for g, quantity := range grades {
		if err := AddGrade(date, code, grade(g), quantity); err != nil {
			return err
		}
	}

	return nil
}

var (
	lastRawIndex   = 2
	lastGradeIndex = 2
)

func AddGrade(date, code string, grade grade, quantity float64) error {
	if quantity == 0 {
		return nil
	}

	const (
		sheet = "Grades os "
	)

	if err := out.DuplicateRow(sheet, lastGradeIndex); err != nil {
		return err
	}
	if err := out.SetCellStr(sheet, fmt.Sprintf("A%d", lastGradeIndex), date); err != nil {
		return err
	}
	if err := out.SetCellStr(sheet, fmt.Sprintf("B%d", lastGradeIndex), code); err != nil {
		return err
	}
	if err := out.SetCellFloat(sheet, fmt.Sprintf("C%d", lastGradeIndex), quantity, 2, 64); err != nil {
		return err
	}
	if err := out.SetCellValue(sheet, fmt.Sprintf("D%d", lastGradeIndex), grade); err != nil {
		return err
	}
	if err := setSupplierCells(sheet, code, lastGradeIndex); err != nil {
		return err
	}

	lastGradeIndex++
	return nil
}

func AddRaw(data, suplier string, quantity float64) error {
	if quantity == 0 {
		return nil
	}

	const (
		sheet = "Raw os"
	)

	if err := out.DuplicateRow(sheet, lastRawIndex); err != nil {
		return err
	}
	if err := out.SetCellStr(sheet, fmt.Sprintf("A%d", lastRawIndex), data); err != nil {
		return err
	}
	if err := out.SetCellStr(sheet, fmt.Sprintf("B%d", lastRawIndex), suplier); err != nil {
		return err
	}
	if err := out.SetCellFloat(sheet, fmt.Sprintf("C%d", lastRawIndex), quantity, 2, 64); err != nil {
		return err
	}
	if err := out.SetCellStr(sheet, fmt.Sprintf("D%d", lastRawIndex), "خام"); err != nil {
		return err
	}
	if err := setSupplierCells(sheet, suplier, lastRawIndex); err != nil {
		return err
	}

	lastRawIndex++
	return nil
}
