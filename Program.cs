using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using Spectre.Console;

class Program
{
    private static readonly Dictionary<string, string> StateMap = new()
    {
        {"01","J&K"}, {"02","HP"}, {"03","Punjab"}, {"04","Chandigarh"}, {"05","Uttarakhand"},
        {"06","Haryana"}, {"07","Delhi"}, {"08","Rajasthan"}, {"09","UP"}, {"10","Bihar"},
        {"11","Sikkim"}, {"12","Arunachal"}, {"13"}, {"14","Manipur"}, {"15","Mizoram"},
        {"16","Tripura"}, {"17","Meghalaya"}, {"18"}, {"19","WB"}, {"20","Jharkhand"},
        {"21","Odisha"}, {"22","Chhattisgarh"}, {"23","MP"}, {"24","Gujarat"}, {"26","DNHDD"},
        {"27","Maharashtra"}, {"29"}, {"30","Goa"}, {"31","Lakshadweep"}, {"32","Kerala"},
        {"33","TN"}, {"34","Puducherry"}, {"35"}, {"36","Telangana"}, {"37","Andhra Pradesh"},
        {"38","Ladakh"}, {"97","UN Bodies"}, {"99","Foreign Entities"}
    };

    static void Main(string[] args)
    {
        string currentFolder = AppDomain.CurrentDomain.BaseDirectory;
        AnsiConsole.Write(new Rule("[yellow]KARZA DEEP-CONSOLIDATION ENGINE v2026.005[/]").LeftAligned());

        try
        {
            // --- PHASE 1: DIRECTORY SCANNING ---
            var files = Directory.GetFiles(currentFolder, "*.xlsx")
                .Where(f => !Path.GetFileName(f).StartsWith("CONSOLIDATED_", StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (!files.Any())
            {
                AnsiConsole.MarkupLine("[red]!! No source Karza reports found in current execution directory.[/]");
                return;
            }

            AnsiConsole.MarkupLine("[yellow][1/3] Mapping files to entity profiles...[/]");
            var fileDataList = new List<KarzaFileProfile>();

            foreach (var file in files)
            {
                using var workbook = new XLWorkbook(file);
                var profileWs = workbook.Worksheet("Entity Profile");
                
                string legalName = profileWs.Cell("B3").GetText().Trim();
                string tradeName = profileWs.Cell("B4").GetText().Trim();
                string pan = profileWs.Cell("B5").GetText().Trim();
                string gstIn = profileWs.Cell("B6").GetText().Trim();

                if (string.IsNullOrEmpty(pan) || pan.Length != 10)
                {
                    if (gstIn.Length >= 15) pan = gstIn.Substring(2, 10);
                }
                if (string.IsNullOrEmpty(pan)) pan = "UNKNOWNPAN";

                string chosenName = tradeName;
                if (string.IsNullOrWhiteSpace(chosenName) || chosenName == "-" || chosenName.Equals("NA", StringComparison.OrdinalIgnoreCase))
                {
                    chosenName = legalName;
                }
                if (string.IsNullOrWhiteSpace(chosenName)) chosenName = "Unknown_Entity";

                string safeName = string.Join("_", chosenName.Split(Path.GetInvalidFileNameChars()));
                string stateCode = gstIn.Length >= 15 ? gstIn.Substring(0, 2) : "00";
                string suffix = gstIn.Length >= 15 ? gstIn.Substring(gstIn.Length - 3, 3) : "XXX";

                fileDataList.Add(new KarzaFileProfile(file, pan, safeName, gstIn, stateCode, suffix));
            }

            // Grouping Logic matching PowerShell AOT specifications
            var entityGroups = fileDataList.GroupBy(f => 
                (f.PAN.Length == 10 && f.PAN[3] == 'P') ? $"{f.PAN}_{f.TradeName}" : f.PAN);

            // --- PHASE 2 & 3: EXTRACTION AND COMPILED WRITE ---
            foreach (var group in entityGroups)
            {
                var groupItems = group.ToList();
                string currentPan = groupItems[0].PAN;
                string currentName = groupItems[0].TradeName;

                AnsiConsole.Write(new Panel(new Markup($"[green]ACTIVE ENTITY: {currentName} ({currentPan})[/]")).Expand());
                
                var stateCounts = groupItems.GroupBy(i => i.StateCode).ToDictionary(g => g.Key, g => g.Count());
                var summaryData = new List<SummaryRecord>();
                var matrixData = new List<MatrixRecord>();
                var relatedPANs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var panToNameMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                AnsiConsole.MarkupLine("[yellow][2/3] Executing in-memory data streaming loops...[/]");

                foreach (var item in groupItems)
                {
                    string stateName = StateMap.TryGetValue(item.StateCode, out var name) ? name : "Unknown";
                    string stateHeader = stateCounts[item.StateCode] > 1 ? $"{item.StateCode}-{stateName}-{item.Suffix}" : $"{item.StateCode}-{stateName}";

                    AnsiConsole.MarkupLine($"      [gray]Streaming Matrix block:[/] {stateHeader} ({Path.GetFileName(item.FilePath)})");
                    using var workbook = new XLWorkbook(item.FilePath);

                    // Extraction Loop 1: Related Parties
                    foreach (var sheetName in new[] { "Related Party Sales - Monthly", "Related Party Purchases-Monthly" })
                    {
                        if (!workbook.TryGetWorksheet(sheetName, out var wsRP)) continue;
                        int rpMaxRow = wsRP.LastRowUsed()?.RowNumber() ?? 100;
                        for (int c = 2; c <= 200; c += 8)
                        {
                            int blankStreak = 0;
                            for (int r = 4; r <= rpMaxRow; r++)
                            {
                                string rpP = wsRP.Cell(r, c).GetText().Trim();
                                if (rpP.Length == 10) { relatedPANs.Add(rpP); blankStreak = 0; }
                                else if (string.IsNullOrEmpty(rpP) && ++blankStreak > 50) break;
                            }
                        }
                    }

                    // Extraction Loop 2: Summary Framework & Fallbacks
                    var gstrWs = workbook.Worksheet("GSTR1 vs 3B");
                    var fileMonths = new Dictionary<string, MonthSummary>(StringComparer.OrdinalIgnoreCase);
                    int gstrMax = gstrWs.LastRowUsed()?.RowNumber() ?? 50;

                    for (int i = 4; i <= gstrMax + 20; i++)
                    {
                        string monthStr = gstrWs.Cell(i, 1).GetText().Trim();
                        if (!IsTextMonthValid(monthStr)) continue;

                        double gi1 = 0, gt1 = 0, gi3b = 0, gt3b = 0;
                        double.TryParse(gstrWs.Cell(i, 2).Value.ToString(), out gi1);
                        double.TryParse(gstrWs.Cell(i, 3).Value.ToString(), out gt1);
                        double.TryParse(gstrWs.Cell(i, 4).Value.ToString(), out gi3b);
                        double.TryParse(gstrWs.Cell(i, 5).Value.ToString(), out gt3b);

                        bool fallbackTriggered = false;
                        double finalGt = gt3b, finalGi = gi3b;

                        if (finalGt == 0 && gt1 > 0) { finalGt = gt1; fallbackTriggered = true; }
                        if (finalGi == 0 && gi1 > 0) { finalGi = gi1; fallbackTriggered = true; }

                        fileMonths[monthStr] = new MonthSummary(finalGt, finalGi, fallbackTriggered);
                    }

                    // Extraction Loop 3: Unlimited Row Deep Scan
                    foreach (var type in new[] { "Customer", "Supplier" })
                    {
                        var ws = workbook.Worksheet($"{type} Wise - Monthly Data");
                        int sheetMaxRow = ws.LastRowUsed()?.RowNumber() ?? 1000;

                        for (int c = 1; c <= 400; c += 9)
                        {
                            string monthHeader = ws.Cell(2, c).GetText().Trim();
                            if (!fileMonths.ContainsKey(monthHeader)) continue;

                            int blankStreak = 0;
                            for (int r = 4; r <= 100000; r++)
                            {
                                string serial = ws.Cell(r, c).GetText().Trim();
                                string cP = ws.Cell(r, c + 1).GetText().Trim();
                                string cN = ws.Cell(r, c + 2).GetText().Trim();

                                if (string.IsNullOrEmpty(serial) && string.IsNullOrEmpty(cP) && string.IsNullOrEmpty(cN))
                                {
                                    if (++blankStreak > 50) break;
                                    continue;
                                }
                                blankStreak = 0;

                                if (serial.Contains("Total", StringComparison.OrdinalIgnoreCase) || 
                                    cP.Contains("Total", StringComparison.OrdinalIgnoreCase) || 
                                    cN.Contains("Total", StringComparison.OrdinalIgnoreCase)) continue;

                                if (string.IsNullOrEmpty(cP)) cP = "UNREGISTERED";
                                if (cP != "UNREGISTERED" && !string.IsNullOrWhiteSpace(cN) && cN != "-") panToNameMap[cP] = cN;

                                double vT = 0, vI = 0;
                                double.TryParse(ws.Cell(r, c + 3).Value.ToString(), out vT);
                                double.TryParse(ws.Cell(r, c + 5).Value.ToString(), out vI);

                                if (vT == 0 && vI == 0) continue;

                                if (cP.Equals(currentPan, StringComparison.OrdinalIgnoreCase))
                                {
                                    if (type == "Customer") { fileMonths[monthHeader].IT_C += vT; fileMonths[monthHeader].II_C += vI; }
                                    else { fileMonths[monthHeader].IT_S += vT; fileMonths[monthHeader].II_S += vI; }
                                }
                                else
                                {
                                    matrixData.Add(new MatrixRecord(cN, cP, stateHeader, monthHeader, vT, vI, type));
                                }
                            }
                        }
                    }

                    foreach (var kp in fileMonths)
                    {
                        summaryData.Add(new SummaryRecord(kp.Key, stateHeader, "Customer", kp.Value.GT, kp.Value.GI, kp.Value.IT_C, kp.Value.II_C, kp.Value.IsFallback));
                        summaryData.Add(new SummaryRecord(kp.Key, stateHeader, "Supplier", kp.Value.GT, kp.Value.GI, kp.Value.IT_S, kp.Value.II_S, kp.Value.IsFallback));
                    }
                }

                // PAN Mapping Correction
                foreach (var mRec in matrixData)
                {
                    if (string.IsNullOrWhiteSpace(mRec.Name) || mRec.Name == "-")
                    {
                        mRec.Name = panToNameMap.TryGetValue(mRec.PAN, out var mappedName) ? mappedName : "Unknown Party";
                    }
                    mRec.IsRelated = relatedPANs.Contains(mRec.PAN);
                }

                // --- PHASE 3: COMPILING EXCEL WORKBOOK ---
                AnsiConsole.MarkupLine("[yellow][3/3] Compiling structures into Open-XML layers...[/]");
                string outPath = Path.Combine(currentFolder, $"CONSOLIDATED_{currentPan}_{currentName}.xlsx");
                
                using var outWb = new XLWorkbook();
                var uniqueMonths = summaryData.Select(s => s.Month).Union(matrixData.Select(m => m.Month))
                    .Distinct().Where(m => !string.IsNullOrEmpty(m))
                    .OrderBy(m => DateTime.ParseExact(m, "MMM-yy", null)).ToList();
                var uniqueStates = summaryData.Select(s => s.StateHeader).Distinct().OrderBy(s => s).ToList();

                // Generate Revenue Sheets
                var netCfgs = new[] {
                    new NetConfig("Tax. Value - Internal Sales", "GT", "IT", "Customer", new[]{"Gross Revenue - Taxable Value", "Internal Sales - Taxable Value", "Net Revenue - Taxable Value"}),
                    new NetConfig("Inv. Value - Internal Sales", "GI", "II", "Customer", new[]{"Gross Revenue - Invoice Value", "Internal Sales - Invoice Value", "Net Revenue - Invoice Value"}),
                    new NetConfig("Tax. Value - Internal Purchases", "GT", "IT", "Supplier", new[]{"Gross Revenue - Taxable Value", "Internal Purchases - Taxable Value", "Net Revenue - Taxable Value"}),
                    new NetConfig("Inv. Value - Internal Purchases", "GI", "II", "Supplier", new[]{"Gross Revenue - Invoice Value", "Internal Purchases - Invoice Value", "Net Revenue - Invoice Value"})
                };

                foreach (var cfg in netCfgs)
                {
                    var ws = outWb.Worksheets.Add(cfg.SheetName);
                    int r = 1;
                    string[] modes = { "Gross", "Internal", "Net" };

                    for (int i = 0; i < 3; i++)
                    {
                        ws.Cell(r, 1).Value = cfg.Labels[i];
                        ws.Cell(r, 1).Style.Font.Bold = true;
                        ws.Cell(r + 1, 1).Value = "Month";
                        ws.Cell(r + 1, 1).Style.DateFormat.SetNumberFormat("@");

                        int cc = 2;
                        foreach (var s in uniqueStates) ws.Cell(r + 1, cc++).Value = s;
                        ws.Cell(r + 1, cc).Value = "Total";

                        int dr = r + 2;
                        foreach (var m in uniqueMonths)
                        {
                            ws.Cell(dr, 1).Value = m;
                            ws.Cell(dr, 1).Style.DateFormat.SetNumberFormat("@");
                            cc = 2;

                            foreach (var s in uniqueStates)
                            {
                                var recs = summaryData.Where(sd => sd.Month == m && sd.StateHeader == s && sd.Type == cfg.TargetType).ToList();
                                double val = 0;
                                bool isFallback = false;

                                if (recs.Any())
                                {
                                    isFallback = recs.Any(rc => rc.IsFallback);
                                    double grossSum = recs.Sum(rc => cfg.GCol == "GT" ? rc.GT : rc.GI);
                                    double internalSum = recs.Sum(rc => cfg.ICol == "IT" ? rc.IT : rc.II);

                                    if (modes[i] == "Gross") val = grossSum;
                                    else if (modes[i] == "Internal") val = internalSum;
                                    else val = grossSum - internalSum;
                                }

                                var cell = ws.Cell(dr, cc++);
                                cell.Value = val;

                                if (isFallback && modes[i] != "Internal")
                                {
                                    cell.Style.Font.Italic = true;
                                    cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#FFFFD0");
                                }
                            }

                            ws.Cell(dr, cc).FormulaA1 = $"=SUM(B{dr}:{GetColLetter(cc - 1)}{dr})";
                            dr++;
                        }
                        r = dr + 2;
                    }
                    ws.Columns().AdjustToContents();
                    ws.SelectedRanges.Clear();
                }

                // Generate Matrix Sheets
                var matrixCfgs = new[] {
                    new MatrixConfig("Detailed_Customer_Taxable", "Customer", "T"),
                    new MatrixConfig("Detailed_Customer_Invoice", "Customer", "I"),
                    new MatrixConfig("Detailed_Supplier_Taxable", "Supplier", "T"),
                    new MatrixConfig("Detailed_Supplier_Invoice", "Supplier", "I")
                };

                foreach (var mCfg in matrixCfgs)
                {
                    var ws = outWb.Worksheets.Add(mCfg.SheetName);
                    ws
