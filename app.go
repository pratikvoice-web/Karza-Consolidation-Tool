package main

import (
	"context"
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"sort"
	"strconv"
	"strings"
	"time"

	"github.com/wailsapp/wails/v2/pkg/runtime"
	"github.com/xuri/excelize/v2"
)

type App struct {
	ctx context.Context
}

func NewApp() *App {
	return &App{}
}

func (a *App) startup(ctx context.Context) {
	a.ctx = ctx
}

type FileMetadata struct {
	FilePath  string
	PAN       string
	TradeName string
	GSTIN     string
	StateCode string
	Suffix    string
}

type MonthData struct {
	GrossTaxable             float64
	GrossInvoice             float64
	InternalTaxableCustomer  float64
	InternalInvoiceCustomer  float64
	InternalTaxableSupplier  float64
	InternalInvoiceSupplier  float64
	IsFallback               bool
}

type SummaryRecord struct {
	Month           string
	State           string
	Type            string
	GrossTaxable    float64
	GrossInvoice    float64
	InternalTaxable float64
	InternalInvoice float64
	IsFallback      bool
}

type MatrixRecord struct {
	Name           string
	PAN            string
	State          string
	Month          string
	Taxable        float64
	Invoice        float64
	IsRelatedParty bool
	Type           string
}

var StateMap = map[string]string{
	"01": "J&K", "02": "HP", "03": "Punjab", "04": "Chandigarh", "05": "Uttarakhand",
	"06": "Haryana", "07": "Delhi", "08": "Rajasthan", "09": "UP", "10": "Bihar",
	"11": "Sikkim", "12": "Arunachal", "13": "Nagaland", "14": "Manipur", "15": "Mizoram",
	"16": "Tripura", "17": "Meghalaya", "18": "Assam", "19": "WB", "20": "Jharkhand",
	"21": "Odisha", "22": "Chhattisgarh", "23": "MP", "24": "Gujarat", "26": "DNHDD",
	"27": "Maharashtra", "29": "Karnataka", "30": "Goa", "31": "Lakshadweep", "32": "Kerala",
	"33": "TN", "34": "Puducherry", "35": "A&N Islands", "36": "Telangana", "37": "Andhra Pradesh",
	"38": "Ladakh", "97": "UN Bodies", "99": "Foreign Entities",
}

func NormalizeEntityName(name string) string {
	if name == "" {
		return "UNKNOWN_ENTITY"
	}
	name = strings.ToUpper(strings.TrimSpace(name))
	rePvt := regexp.MustCompile(`\b(?:PRIVATE|PVT\.?|\(P\))\s*(?:LIMITED|LTD\.?)\b`)
	name = rePvt.ReplaceAllString(name, "PVT LTD")
	reLtd := regexp.MustCompile(`\b(?:LIMITED|LTD\.?)\b`)
	name = reLtd.ReplaceAllString(name, "LTD")
	reSpace := regexp.MustCompile(`\s+`)
	name = reSpace.ReplaceAllString(name, " ")
	reClean := regexp.MustCompile(`[\\/:*?"<>|]`)
	name = reClean.ReplaceAllString(name, "_")
	return strings.TrimSpace(name)
}

func GetFinancialYear(mmmYY string) string {
	t, err := time.Parse("Jan-06", mmmYY)
	if err != nil {
		return "FY_UNKNOWN"
	}
	year := t.Year()
	if t.Month() >= 4 {
		return fmt.Sprintf("FY%02d-%02d", year%100, (year+1)%100)
	}
	return fmt.Sprintf("FY%02d-%02d", (year-1)%100, year%100)
}

func SafeFloat(val string) float64 {
	val = strings.TrimSpace(val)
	if val == "" || val == "-" {
		return 0
	}
	f, err := strconv.ParseFloat(val, 64)
	if err != nil {
		return 0
	}
	return f
}

func GetColumnLetter(col int) string {
	letter := ""
	for col > 0 {
		mod := (col - 1) % 26
		letter = string(rune(65+mod)) + letter
		col = (col - mod) / 26
	}
	return letter
}

func (a *App) SelectDirectory() string {
	res, err := runtime.OpenDirectoryDialog(a.ctx, runtime.OpenDialogOptions{
		Title: "Select Target Operation Directory Base",
	})
	if err != nil {
		return ""
	}
	return res
}

func (a *App) ExecuteConsolidation(inputFolder, outputFolder string) string {
	runtime.EventsEmit(a.ctx, "log", "Initializing scanning matrix on filesystem elements...")
	
	files, err := os.ReadDir(inputFolder)
	if err != nil {
		return fmt.Sprintf("Directory acquisition error: %s", err.Error())
	}

	var excelFiles []string
	for _, f := range files {
		if !f.IsDir() && strings.HasSuffix(strings.ToLower(f.Name()), ".xlsx") && !strings.HasPrefix(strings.ToUpper(f.Name()), "CONSOLIDATED_") {
			excelFiles = append(excelFiles, filepath.Join(inputFolder, f.Name()))
		}
	}

	if len(excelFiles) == 0 {
		return "No target Karza source ledger files discovered inside target directory."
	}

	var fileDataList []FileMetadata
	for _, path := range excelFiles {
		f, err := excelize.OpenFile(path)
		if err != nil {
			continue
		}
		
		b3, _ := f.GetCellValue("Entity Profile", "B3")
		b4, _ := f.GetCellValue("Entity Profile", "B4")
		b5, _ := f.GetCellValue("Entity Profile", "B5")
		b6, _ := f.GetCellValue("Entity Profile", "B6")
		_ = f.Close()

		pan := strings.TrimSpace(b5)
		if len(pan) != 10 && len(b6) >= 15 {
			pan = b6[2:12]
		}
		if pan == "" {
			pan = "UNKNOWNPAN"
		}

		name := b4
		if name == "" || name == "-" || name == "NA" {
			name = b3
		}
		safeName := NormalizeEntityName(name)

		stateCode := "00"
		if len(b6) >= 15 {
			stateCode = b6[0:2]
		}
		suffix := "XXX"
		if len(b6) >= 15 {
			suffix = b6[len(b6)-3:]
		}

		fileDataList = append(fileDataList, FileMetadata{
			FilePath:  path,
			PAN:       pan,
			TradeName: safeName,
			GSTIN:     b6,
			StateCode: stateCode,
			Suffix:    suffix,
		})
	}

	entityGroups := make(map[string][]FileMetadata)
	for _, fd := range fileDataList {
		key := fd.PAN
		if len(fd.PAN) == 10 && fd.PAN[3] == 'P' {
			key = fmt.Sprintf("%s_%s", fd.PAN, fd.TradeName)
		}
		entityGroups[key] = append(entityGroups[key], fd)
	}

	for _, items := range entityGroups {
		currentPan := items[0].PAN
		currentName := items[0].TradeName

		runtime.EventsEmit(a.ctx, "log", fmt.Sprintf("Processing Profile Boundaries for Entity: %s", currentName))

		stateCounts := make(map[string]int)
		for _, item := range items {
			stateCounts[item.StateCode]++
		}

		var summaryData []SummaryRecord
		var matrixData []MatrixRecord
		relatedPANs := make(map[string]bool)
		panToNameMap := make(map[string]string)

		for idx, item := range items {
			stateName, exists := StateMap[item.StateCode]
			if !exists {
				stateName = "Unknown"
			}
			stHead := fmt.Sprintf("%s-%s", item.StateCode, stateName)
			if stateCounts[item.StateCode] > 1 {
				stHead = fmt.Sprintf("%s-%s-%s", item.StateCode, stateName, item.Suffix)
			}

			pct := (float64(idx+1) / float64(len(items))) * 100
			runtime.EventsEmit(a.ctx, "extract", map[string]interface{}{"val": pct, "txt": fmt.Sprintf("Extracting Layer (%d/%d): %s", idx+1, len(items), stHead)})

			wb, err := excelize.OpenFile(item.FilePath)
			if err != nil {
				continue
			}

			for _, sName := range []string{"Related Party Sales - Monthly", "Related Party Purchases-Monthly"} {
				rows, err := wb.GetRows(sName)
				if err == nil && len(rows) > 3 {
					for c := 1; c < len(rows[3]); c += 8 {
						blankStreak := 0
						for r := 3; r < len(rows); r++ {
							if c < len(rows[r]) {
								rpp := strings.TrimSpace(rows[r][c])
								if len(rpp) == 10 {
									relatedPANs[rpp] = true
									blankStreak = 0
								} else {
									blankStreak++
									if blankStreak > 50 {
										break
									}
								}
							}
						}
					}
				}
			}

			fileMonths := make(map[string]*MonthData)
			gRows, err := wb.GetRows("GSTR1 vs 3B")
			if err == nil {
				reM := regexp.MustCompile(`^[A-Za-z]{3}-\d{2}$`)
				for r := 3; r < len(gRows); r++ {
					if len(gRows[r]) > 0 {
						m := strings.TrimSpace(gRows[r][0])
						if reM.MatchString(m) {
							gi1 := 0.0; gt1 := 0.0; gi3b := 0.0; gt3b := 0.0
							if len(gRows[r]) > 1 { gi1 = SafeFloat(gRows[r][1]) }
							if len(gRows[r]) > 2 { gt1 = SafeFloat(gRows[r][2]) }
							if len(gRows[r]) > 3 { gi3b = SafeFloat(gRows[r][3]) }
							if len(gRows[r]) > 4 { gt3b = SafeFloat(gRows[r][4]) }

							fallback := false
							if gt3b == 0 && gt1 > 0 { gt3b = gt1; fallback = true }
							if gi3b == 0 && gi1 > 0 { gi3b = gi1; fallback = true }

							fileMonths[m] = &MonthData{GrossTaxable: gt3b, GrossInvoice: gi3b, IsFallback: fallback}
						}
					}
				}
			}

			for _, mType := range []string{"Customer", "Supplier"} {
				sName := fmt.Sprintf("%s Wise - Monthly Data", mType)
				mRows, err := wb.GetRows(sName)
				if err == nil && len(mRows) > 2 {
					for c := 0; c < len(mRows[1]); c += 9 {
						m := strings.TrimSpace(mRows[1][c])
						md, exists := fileMonths[m]
						if !exists {
							continue
						}

						blankStreak := 0
						for r := 3; r < len(mRows); r++ {
							if c+2 >= len(mRows[r]) {
								blankStreak++
								if blankStreak > 50 { break }
								continue
							}
							serial := strings.TrimSpace(mRows[r][c])
							cp := strings.TrimSpace(mRows[r][c+1])
							cn := strings.TrimSpace(mRows[r][c+2])

							if serial == "" && cp == "" && cn == "" {
								blankStreak++
								if blankStreak > 50 { break }
								continue
							}
							blankStreak = 0

							if strings.Contains(strings.ToLower(serial), "total") || strings.Contains(strings.ToLower(cp), "total") || strings.Contains(strings.ToLower(cn), "total") {
								continue
							}

							if cp == "" { cp = "UNREGISTERED" }
							var normCn string
							if cp != "UNREGISTERED" && cn != "" && cn != "-" {
								normCn = NormalizeEntityName(cn)
								panToNameMap[cp] = normCn
							}

							vt := 0.0; vi := 0.0
							if c+3 < len(mRows[r]) { vt = SafeFloat(mRows[r][c+3]) }
							if c+5 < len(mRows[r]) { vi = SafeFloat(mRows[r][c+5]) }

							if vt == 0 && vi == 0 { continue }

							if cp == currentPan {
								if mType == "Customer" {
									md.InternalTaxableCustomer += vt
									md.InternalInvoiceCustomer += vi
								} else {
									md.InternalTaxableSupplier += vt
									md.InternalInvoiceSupplier += vi
								}
							} else {
								matrixData = append(matrixData, MatrixRecord{
									Name: normCn, PAN: cp, State: stHead, Month: m, Taxable: vt, Invoice: vi, Type: mType,
								})
							}
						}
					}
				}
			}
			_ = wb.Close()

			for m, d := range fileMonths {
				summaryData = append(summaryData, SummaryRecord{Month: m, State: stHead, Type: "Customer", GrossTaxable: d.GrossTaxable, GrossInvoice: d.GrossInvoice, InternalTaxable: d.InternalTaxableCustomer, InternalInvoice: d.InternalInvoiceCustomer, IsFallback: d.IsFallback})
				summaryData = append(summaryData, SummaryRecord{Month: m, State: stHead, Type: "Supplier", GrossTaxable: d.GrossTaxable, GrossInvoice: d.GrossInvoice, InternalTaxable: d.InternalTaxableSupplier, InternalInvoice: d.InternalInvoiceSupplier, IsFallback: d.IsFallback})
			}
		}
		runtime.EventsEmit(a.ctx, "extract", map[string]interface{}{"val": 100.0, "txt": "Ledger extraction pass completed safely."})

		for idx, md := range matrixData {
			matrixData[idx].IsRelatedParty = relatedPANs[md.PAN]
			if md.Name == "" || md.Name == "-" {
				if mappedName, ok := panToNameMap[md.PAN]; ok {
					matrixData[idx].Name = mappedName
				} else if md.PAN == "UNREGISTERED" {
					matrixData[idx].Name = "CONSUMER / UNREGISTERED SALES"
				} else {
					matrixData[idx].Name = "UNKNOWN COUNTERPARTY"
				}
			}
		}

		monthSet := make(map[string]bool)
		for _, s := range summaryData { monthSet[s.Month] = true }
		for _, m := range matrixData { monthSet[m.Month] = true }
		var uniqueMonths []string
		for m := range monthSet { if m != "" { uniqueMonths = append(uniqueMonths, m) } }
		sort.Slice(uniqueMonths, func(i, j int) bool {
			ti, _ := time.Parse("Jan-02", uniqueMonths[i])
			tj, _ := time.Parse("Jan-02", uniqueMonths[j])
			return ti.Before(tj)
		})

		stateSet := make(map[string]bool)
		for _, s := range summaryData { stateSet[s.State] = true }
		var uniqueStates []string
		for st := range stateSet { uniqueStates = append(uniqueStates, st) }
		sort.Strings(uniqueStates)

		fyMap := make(map[string][]string)
		for _, m := range uniqueMonths {
			fy := GetFinancialYear(m)
			fyMap[fy] = append(fyMap[fy], m)
		}
		var uniqueFYs []string
		for fy := range fyMap { uniqueFYs = append(uniqueFYs, fy) }
		sort.Strings(uniqueFYs)

		outWb := excelize.NewFile()
		_ = outWb.DeleteSheet("Sheet1")

		_, _ = outWb.NewSheet("Index")
		_ = outWb.SetCellValue("Index", "A1", "Consolidated GST Karza")
		_ = outWb.SetCellValue("Index", "A3", "Entity Name:")
		_ = outWb.SetCellValue("Index", "B3", currentName)
		_ = outWb.SetCellValue("Index", "A4", "PAN:")
		_ = outWb.SetCellValue("Index", "B4", currentPan)
		_ = outWb.SetCellValue("Index", "A6", "Table of Contents")
		
		indexRow := 8
		sheetCount := 1

		addToIndex := func(sName, desc string) {
			_ = outWb.SetCellValue("Index", fmt.Sprintf("A%d", indexRow), sheetCount)
			_ = outWb.SetCellValue("Index", fmt.Sprintf("B%d", indexRow), sName)
			_ = outWb.SetCellHyperLink("Index", fmt.Sprintf("B%d", indexRow), fmt.Sprintf("'%s'!A1", sName), "Location")
			_ = outWb.SetCellValue("Index", fmt.Sprintf("C%d", indexRow), desc)
			indexRow++
			sheetCount++
		}

		netConfigs := []struct {
			SheetName string; Target string; IsTax bool; Labels []string
		}{
			{"Tax. Value - Internal Sales", "Customer", true, []string{"Gross Revenue - Taxable Value", "Internal Sales - Taxable Value", "Net Revenue - Taxable Value"}},
			{"Inv. Value - Internal Sales", "Customer", false, []string{"Gross Revenue - Invoice Value", "Internal Sales - Invoice Value", "Net Revenue - Invoice Value"}},
			{"Tax. Value - Internal Purchases", "Supplier", true, []string{"Gross Revenue - Taxable Value", "Internal Purchases - Taxable Value", "Net Revenue - Taxable Value"}},
			{"Inv. Value - Internal Purchases", "Supplier", false, []string{"Gross Revenue - Invoice Value", "Internal Purchases - Invoice Value", "Net Revenue - Invoice Value"}},
		}

		numStyle, _ := outWb.NewStyle(&excelize.Style{CustomNumFmt: &[]string{"#,##0.00"}[0]})

		for cIdx, cfg := range netConfigs {
			runtime.EventsEmit(a.ctx, "compile", map[string]interface{}{"val": (float64(cIdx+1) / 10.0) * 100, "txt": fmt.Sprintf("Writing Array Map: %s", cfg.SheetName)})
			_, _ = outWb.NewSheet(cfg.SheetName)
			addToIndex(cfg.SheetName, fmt.Sprintf("Monthly Summary - %s", cfg.SheetName))

			rowTracker := 1
			for _, block := range []string{"Gross", "Internal", "Net"} {
				lbl := cfg.Labels[0]
				if block == "Internal" { lbl = cfg.Labels[1] } else if block == "Net" { lbl = cfg.Labels[2] }
				
				_ = outWb.SetCellValue(cfg.SheetName, fmt.Sprintf("A%d", rowTracker), lbl)
				_ = outWb.SetCellValue(cfg.SheetName, fmt.Sprintf("A%d", rowTracker+1), "Financial Year / Month")
				
				cCol := 2
				for _, st := range uniqueStates {
					_ = outWb.SetCellValue(cfg.SheetName, fmt.Sprintf("%s%d", GetColumnLetter(cCol), rowTracker+1), st)
					cCol++
				}
				_ = outWb.SetCellValue(cfg.SheetName, fmt.Sprintf("%s%d", GetColumnLetter(cCol), rowTracker+1), "Total")

				dataRow := rowTracker + 2
				for _, fy := range uniqueFYs {
					fyRow := dataRow
					_ = outWb.SetCellValue(cfg.SheetName, fmt.Sprintf("A%d", fyRow), fy)
					dataRow++
					startGroup := dataRow

					for _, m := range fyMap[fy] {
						_ = outWb.SetCellValue(cfg.SheetName, fmt.Sprintf("A%d", dataRow), m)
						
						colSub := 2
						for _, st := range uniqueStates {
							cellVal := 0.0
							for _, s := range summaryData {
								if s.Month == m && s.State == st && s.Type == cfg.Target {
									if block == "Gross" { if cfg.IsTax { cellVal += s.GrossTaxable } else { cellVal += s.GrossInvoice } }
									if block == "Internal" { if cfg.IsTax { cellVal += s.InternalTaxable } else { cellVal += s.InternalInvoice } }
									if block == "Net" { if cfg.IsTax { cellVal += (s.GrossTaxable - s.InternalTaxable) } else { cellVal += (s.GrossInvoice - s.InternalInvoice) } }
								}
							}
							cellRef := fmt.Sprintf("%s%d", GetColumnLetter(colSub), dataRow)
							_ = outWb.SetCellValue(cfg.SheetName, cellRef, cellVal)
							_ = outWb.SetCellStyle(cfg.SheetName, cellRef, cellRef, numStyle)
							colSub++
						}
						
						rowTotCell := fmt.Sprintf("%s%d", GetColumnLetter(colSub), dataRow)
						_ = outWb.SetCellFormula(cfg.SheetName, rowTotCell, fmt.Sprintf("=SUM(B%d:%s%d)", dataRow, GetColumnLetter(colSub-1), dataRow))
						_ = outWb.SetCellStyle(cfg.SheetName, rowTotCell, rowTotCell, numStyle)
						dataRow++
					}

					if dataRow > startGroup {
						for r := startGroup; r <= dataRow-1; r++ {
							_ = outWb.SetRowOutlineLevel(cfg.SheetName, r, 1)
						}
						for c := 2; c < colSub; c++ {
							cL := GetColumnLetter(c)
							targetCell := fmt.Sprintf("%s%d", cL, fyRow)
							_ = outWb.SetCellFormula(cfg.SheetName, targetCell, fmt.Sprintf("=SUM(%s%d:%s%d)", cL, startGroup, cL, dataRow-1))
							_ = outWb.SetCellStyle(cfg.SheetName, targetCell, targetCell, numStyle)
						}
					}
				}
				rowTracker = dataRow + 2
			}
		}

		matrixConfigs := []struct{ SheetName, Target, ValType string }{
			{"Detailed_Customer_Taxable", "Customer", "T"},
			{"Detailed_Customer_Invoice", "Customer", "I"},
			{"Detailed_Supplier_Taxable", "Supplier", "T"},
			{"Detailed_Supplier_Invoice", "Supplier", "I"},
		}

		for mIdx, mCfg := range matrixConfigs {
			runtime.EventsEmit(a.ctx, "compile", map[string]interface{}{"val": (float64(mIdx+5) / 10.0) * 100, "txt": fmt.Sprintf("Writing Subledger Layout: %s", mCfg.SheetName)})
			_, _ = outWb.NewSheet(mCfg.SheetName)
			addToIndex(mCfg.SheetName, fmt.Sprintf("Detailed Party-wise Subledger (%s)", mCfg.Target))

			_ = outWb.SetCellValue(mCfg.SheetName, "A1", "Financial Year")
			_ = outWb.SetCellValue(mCfg.SheetName, "A2", "Party / State")
			_ = outWb.SetCellValue(mCfg.SheetName, "B2", "PAN")

			colIdx := 3
			fyTotalCols := make(map[string]string)
			for _, fy := range uniqueFYs {
				startCol := colIdx
				for _, m := range fyMap[fy] {
					_ = outWb.SetCellValue(mCfg.SheetName, fmt.Sprintf("%s2", GetColumnLetter(colIdx)), m)
					colIdx++
				}
				
				fyTotL := GetColumnLetter(colIdx)
				_ = outWb.SetCellValue(mCfg.SheetName, fmt.Sprintf("%s2", fyTotL), fmt.Sprintf("%s Total", fy))
				fyTotalCols[fy] = fyTotL
				
				_ = outWb.SetCellValue(mCfg.SheetName, fmt.Sprintf("%s1", GetColumnLetter(startCol)), fy)
				_ = outWb.MergeCell(mCfg.SheetName, fmt.Sprintf("%s1", GetColumnLetter(startCol)), fmt.Sprintf("%s1", GetColumnLetter(colIdx)))
				
				for c := startCol; c < colIdx; c++ {
					_ = outWb.SetColOutlineLevel(mCfg.SheetName, GetColumnLetter(c), 1)
				}
				colIdx++
			}
			_ = outWb.SetCellValue(mCfg.SheetName, fmt.Sprintf("%s2", GetColumnLetter(colIdx)), "Grand Total")

			panGroups := make(map[string][]MatrixRecord)
			for _, m := range matrixData {
				if m.Type == mCfg.Target { panGroups[m.PAN] = append(panGroups[m.PAN], m) }
			}

			type RankedPan struct { Key string; Total float64 }
			var rankedList []RankedPan
			for k, v := range panGroups {
				tot := 0.0
				for _, r := range v { if mCfg.ValType == "T" { tot += r.Taxable } else { tot += r.Invoice } }
				rankedList = append(rankedList, RankedPan{Key: k, Total: tot})
			}
			sort.Slice(rankedList, func(i, j int) bool { return rankedList[i].Total > rankedList[j].Total })

			dataRow := 3
			for _, rp := range rankedList {
				pItems := panGroups[rp.Key]
				first := pItems[0]

				_ = outWb.SetCellValue(mCfg.SheetName, fmt.Sprintf("A%d", dataRow), first.Name)
				_ = outWb.SetCellValue(mCfg.SheetName, fmt.Sprintf("B%d", dataRow), rp.Key)

				parentRow := dataRow
				dataRow++

				stateGroupMap := make(map[string][]MatrixRecord)
				for _, p := range pItems { stateGroupMap[p.State] = append(stateGroupMap[p.State], p) }

				for st, stItems := range stateGroupMap {
					_ = outWb.SetCellValue(mCfg.SheetName, fmt.Sprintf("A%d", dataRow), fmt.Sprintf("   >> %s", st))
					
					cCol := 3
					var crossFyFormula []string
					for _, fy := range uniqueFYs {
						sCol := cCol
						for _, m := range fyMap[fy] {
							vSum := 0.0
							for _, item := range stItems { if item.Month == m { if mCfg.ValType == "T" { vSum += item.Taxable } else { vSum += item.Invoice } } }
							if vSum > 0 {
								cellRef := fmt.Sprintf("%s%d", GetColumnLetter(cCol), dataRow)
								_ = outWb.SetCellValue(mCfg.SheetName, cellRef, vSum)
								_ = outWb.SetCellStyle(mCfg.SheetName, cellRef, cellRef, numStyle)
							}
							cCol++
						}
						fyCellRef := fmt.Sprintf("%s%d", GetColumnLetter(cCol), dataRow)
						_ = outWb.SetCellFormula(mCfg.SheetName, fyCellRef, fmt.Sprintf("=SUM(%s%d:%s%d)", GetColumnLetter(sCol), dataRow, GetColumnLetter(cCol-1), dataRow))
						_ = outWb.SetCellStyle(mCfg.SheetName, fyCellRef, fyCellRef, numStyle)
						crossFyFormula = append(crossFyFormula, fyCellRef)
						cCol++
					}
					
					grandCellRef := fmt.Sprintf("%s%d", GetColumnLetter(cCol), dataRow)
					_ = outWb.SetCellFormula(mCfg.SheetName, grandCellRef, "="+strings.Join(crossFyFormula, "+"))
					_ = outWb.SetCellStyle(mCfg.SheetName, grandCellRef, grandCellRef, numStyle)
					dataRow++
				}
				
				for r := parentRow + 1; r <= dataRow-1; r++ {
					_ = outWb.SetRowOutlineLevel(mCfg.SheetName, r, 1)
				}
				
				for c := 3; c <= colIdx; c++ {
					colL := GetColumnLetter(c)
					parentCell := fmt.Sprintf("%s%d", colL, parentRow)
					_ = outWb.SetCellFormula(mCfg.SheetName, parentCell, fmt.Sprintf("=SUM(%s%d:%s%d)", colL, parentRow+1, colL, dataRow-1))
					_ = outWb.SetCellStyle(mCfg.SheetName, parentCell, parentCell, numStyle)
				}
			}
		}

		runtime.EventsEmit(a.ctx, "compile", map[string]interface{}{"val": 100.0, "txt": "Finalizing ledger metadata profiles..."})
		
		_, _ = outWb.NewSheet("Audit_Glossary")
		addToIndex("Audit_Glossary", "Reporting Ledger Color Key & System Glossary")

		boldStyle, _ := outWb.NewStyle(&excelize.Style{Font: &excelize.Font{Bold: true}})
		fillYellow, _ := outWb.NewStyle(&excelize.Style{Fill: excelize.Fill{Type: "pattern", Color: []string{"#FFF2CC"}, Pattern: 1}})
		fillGray, _ := outWb.NewStyle(&excelize.Style{Fill: excelize.Fill{Type: "pattern", Color: []string{"#E1E1E1"}, Pattern: 1}})
		italicStyle, _ := outWb.NewStyle(&excelize.Style{Font: &excelize.Font{Italic: true}, Fill: excelize.Fill{Type: "pattern", Color: []string{"#FFF2CC"}, Pattern: 1}})

		_ = outWb.SetCellValue("Audit_Glossary", "A1", "Reporting Ledger Color Key")
		_ = outWb.SetCellStyle("Audit_Glossary", "A1", "A1", boldStyle)
		
		_ = outWb.SetCellStyle("Audit_Glossary", "A3", "A3", fillYellow)
		_ = outWb.SetCellValue("Audit_Glossary", "B3", "Related Party Configuration / Subledger Identifiers")
		
		_ = outWb.SetCellStyle("Audit_Glossary", "A4", "A4", fillGray)
		_ = outWb.SetCellValue("Audit_Glossary", "B4", "Third-Party Verified Operational Vectors")
		
		_ = outWb.SetCellStyle("Audit_Glossary", "A5", "A5", italicStyle)
		_ = outWb.SetCellValue("Audit_Glossary", "B5", "GSTR1 Operational Fallback Values (Triggered when explicit GSTR3B data is filed as missing or 0)")

		_ = outWb.SetColWidth("Audit_Glossary", "A", "A", 5)
		_ = outWb.SetColWidth("Audit_Glossary", "B", "B", 100)
		_ = outWb.SetColWidth("Index", "A", "C", 40)

		outputName := fmt.Sprintf("CONSOLIDATED_%s_%s.xlsx", currentPan, currentName)
		finalPath := filepath.Join(outputFolder, outputName)
		
		if err := outWb.SaveAs(finalPath); err != nil {
			return fmt.Sprintf("Save failed: %v", err)
		}
		
		runtime.EventsEmit(a.ctx, "log", fmt.Sprintf("Export Complete: %s", outputName))
	}
	return "SUCCESS"
}
