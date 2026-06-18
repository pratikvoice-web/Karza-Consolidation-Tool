using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using Spectre.Console;

namespace KarzaConsolidator
{
    class Program
    {
        private static readonly Dictionary<string, string> StateMap = new()
        {
            { "01", "J&K" }, { "02", "HP" }, { "03", "Punjab" }, { "04", "Chandigarh" }, { "05", "Uttarakhand" },
            { "06", "Haryana" }, { "07", "Delhi" }, { "08", "Rajasthan" }, { "09", "UP" }, { "10", "Bihar" },
            { "11", "Sikkim" }, { "12", "Arunachal" }, { "13", "Nagaland" }, { "14", "Manipur" }, { "15", "Mizoram" },
            { "16", "Tripura" }, { "17", "Meghalaya" }, { "18", "Assam" }, { "19", "WB" }, { "20", "Jharkhand" },
            { "21", "Odisha" }, { "22", "Chhattisgarh" }, { "23", "MP" }, { "24", "Gujarat" }, { "26", "DNHDD" },
            { "27", "Maharashtra" }, { "29", "Karnataka" }, { "30", "Goa" }, { "31", "Lakshadweep" }, { "32", "Kerala" },
            { "33", "TN" }, { "34", "Puducherry" }, { "35", "A&N Islands" }, { "36", "Telangana" }, { "37", "Andhra Pradesh" },
            { "38", "Ladakh" }, { "97", "UN Bodies" }, { "99", "Foreign Entities" }
        };

        static void Main(string[] args)
        {
            AnsiConsole.Write(new FigletText("KARZA CORE").Color(Color.FromConsoleColor(ConsoleColor.Cyan)));
            string currentFolder = AppDomain.CurrentDomain.BaseDirectory;

            var excelFiles = Directory.GetFiles(currentFolder, "*.xlsx")
                .Where(f => !Path.GetFileName(f).StartsWith("CONSOLIDATED_"))
                .ToList();

            if (excelFiles.Count == 0)
            {
                AnsiConsole.MarkupLine("[red]!! No target Karza reports found in execution directory.[/]");
                Console.WriteLine("\nPress any key to close...");
                Console.ReadKey();
                return;
            }

            // ESCAPED BRACKETS FOR SPECTRE
            AnsiConsole.MarkupLine($"[yellow][[1/3]][/] Initializing metadata index pass over {excelFiles.Count} workbooks...");
            var fileDataList = new List<FileMetadata>();

            foreach (var file in excelFiles)
            {
                try
                {
                    using var workbook = new XLWorkbook(file);
                    var ws = workbook.Worksheet("Entity Profile");
                    string b3 = ws.Cell("B3").Value.ToString().Trim();
                    string b4 = ws.Cell("B4").Value.ToString().Trim();
                    string b5 = ws.Cell("B5").Value.ToString().Trim();
                    string b6 = ws.Cell("B6").Value.ToString().Trim();

                    string pan = b5;
                    if (string.IsNullOrEmpty(pan) || pan.Length != 10)
                    {
                        if (b6.Length >= 15) pan = b6.Substring(2, 10);
                    }
                    if (string.IsNullOrEmpty(pan)) pan = "UNKNOWNPAN";

                    string name = b4;
                    if (string.IsNullOrWhiteSpace(name) || name == "-" || name == "NA") name = b3;
                    if (string.IsNullOrWhiteSpace(name)) name = "Unknown_Entity";
                    string safeName = Regex.Replace(name, @"[\\/:*?""<>|]", "_");

                    string stateCode = b6.Length >= 15 ? b6.Substring(0, 2) : "00";
                    string suffix = b6.Length >= 15 ? b6.Substring(b6.Length - 3, 3) : "XXX";

                    fileDataList.Add(new FileMetadata(file, pan, safeName, b6, stateCode, suffix));
                }
                catch (Exception ex)
                {
                    AnsiConsole.MarkupLine($"[red]Error parsing profile metadata for file {Markup.Escape(Path.GetFileName(file))}: {Markup.Escape(ex.Message)}[/]");
                }
            }

            var entityGroups = new Dictionary<string, List<FileMetadata>>();
            foreach (var fd in fileDataList)
            {
                string key = fd.PAN;
                if (fd.PAN.Length == 10 && Char.ToUpper(fd.PAN[3]) == 'P')
                {
                    key = $"{fd.PAN}_{fd.TradeName}";
                }
                if (!entityGroups.ContainsKey(key)) entityGroups[key] = new List<FileMetadata>();
                entityGroups[key].Add(fd);
            }

            foreach (var group in entityGroups)
            {
                var items = group.Value;
                string currentPan = items[0].PAN;
                string currentName = items[0].TradeName;

                AnsiConsole.Write(new Rule($"[green]ACTIVE ENTITY: {Markup.Escape(currentName)} ({Markup.Escape(currentPan)})[/]").LeftJustified());

                var stateCounts = items.GroupBy(i => i.StateCode).ToDictionary(g => g.Key, g => g.Count());
                var summaryData = new List<SummaryRecord>();
                var matrixData = new List<MatrixRecord>();
                var relatedPANs = new HashSet<string>();
                var panToNameMap = new Dictionary<string, string>();

                // ESCAPED BRACKETS FOR SPECTRE
                AnsiConsole.MarkupLine("[yellow][[2/3]][/] Running stream processing matrix layers...");

                int fileIndex = 0;
                foreach (var item in items)
                {
                    fileIndex++;
                    string stateName = StateMap.ContainsKey(item.StateCode) ? StateMap[item.StateCode] : "Unknown";
                    string stHead = stateCounts[item.StateCode] > 1 ? $"{item.StateCode}-{stateName}-{item.Suffix}" : $"{item.StateCode}-{stateName}";

                    AnsiConsole.MarkupLine($"      [gray]READING Layer ({fileIndex}/{items.Count}): {Markup.Escape(stHead)}[/]");

                    using var wb = new XLWorkbook(item.FilePath);

                    foreach (var sheetName in new[] { "Related Party Sales - Monthly", "Related Party Purchases-Monthly" })
                    {
                        IXLWorksheet wsRP;
                        if (wb.TryGetWorksheet(sheetName, out wsRP))
                        {
                            int rpMax = wsRP.LastRowUsed()?.RowNumber() ?? 100;
                            for (int c = 2; c <= 200; c += 8)
                            {
                                int blankStreak = 0;
                                for (int r = 4; r <= rpMax + 50; r++)
                                {
                                    string rpp = wsRP.Cell(r, c).Value.ToString().Trim();
                                    if (rpp.Length == 10)
                                    {
                                        relatedPANs.Add(rpp);
                                        blankStreak = 0;
                                    }
                                    else
                                    {
                                        blankStreak++;
                                        if (blankStreak > 50) break;
                                    }
                                }
                            }
                        }
                    }

                    var fileMonths = new Dictionary<string, MonthData>();
                    IXLWorksheet wsG;
                    if (wb.TryGetWorksheet("GSTR1 vs 3B", out wsG))
                    {
                        int gMax = wsG.LastRowUsed()?.RowNumber() ?? 50;
                        for (int i = 4; i <= gMax + 20; i++)
                        {
                            string m = wsG.Cell(i, 1).Value.ToString().Trim();
                            if (Regex.IsMatch(m, @"^[A-Za-z]{3}-\d{2}$"))
                            {
                                double gi1 = SafeDouble(wsG.Cell(i, 2).Value);
                                double gt1 = SafeDouble(wsG.Cell(i, 3).Value);
                                double gi3b = SafeDouble(wsG.Cell(i, 4).Value);
                                double gt3b = SafeDouble(wsG.Cell(i, 5).Value);

                                bool fallback = false;
                                double finalGt = gt3b;
                                double finalGi = gi3b;

                                if (finalGt == 0 && gt1 > 0) { finalGt = gt1; fallback = true; }
                                if (finalGi == 0 && gi1 > 0) { finalGi = gi1; fallback = true; }

                                fileMonths[m] = new MonthData(finalGt, finalGi, fallback);
                            }
                        }
                    }

                    foreach (var type in new[] { "Customer", "Supplier" })
                    {
                        IXLWorksheet wsM;
                        if (wb.TryGetWorksheet($"{type} Wise - Monthly Data", out wsM))
                        {
                            int mMax = wsM.LastRowUsed()?.RowNumber() ?? 1000;
                            for (int c = 1; c <= 400; c += 9)
                            {
                                string m = wsM.Cell(2, c).Value.ToString().Trim();
                                if (!fileMonths.ContainsKey(m)) continue;

                                int blankStreak = 0;
                                for (int r = 4; r <= mMax + 100; r++)
                                {
                                    string serial = wsM.Cell(r, c).Value.ToString().Trim();
                                    string cp = wsM.Cell(r, c + 1).Value.ToString().Trim();
                                    string cn = wsM.Cell(r, c + 2).Value.ToString().Trim();

                                    if (string.IsNullOrEmpty(serial) && string.IsNullOrEmpty(cp) && string.IsNullOrEmpty(cn))
                                    {
                                        blankStreak++;
                                        if (blankStreak > 50) break;
                                        continue;
                                    }
                                    blankStreak = 0;

                                    if (serial.Contains("Total", StringComparison.OrdinalIgnoreCase) ||
                                        cp.Contains("Total", StringComparison.OrdinalIgnoreCase) ||
                                        cn.Contains("Total", StringComparison.OrdinalIgnoreCase)) continue;

                                    if (string.IsNullOrEmpty(cp)) cp = "UNREGISTERED";

                                    if (cp != "UNREGISTERED" && !string.IsNullOrWhiteSpace(cn) && cn != "-")
                                    {
                                        panToNameMap[cp] = cn;
                                    }

                                    double vt = SafeDouble(wsM.Cell(r, c + 3).Value);
                                    double vi = SafeDouble(wsM.Cell(r, c + 5).Value);

                                    if (vt == 0 && vi == 0) continue;

                                    if (cp == currentPan)
                                    {
                                        if (type == "Customer") { fileMonths[m].InternalTaxableCustomer += vt; fileMonths[m].InternalInvoiceCustomer += vi; }
                                        else { fileMonths[m].InternalTaxableSupplier += vt; fileMonths[m].InternalInvoiceSupplier += vi; }
                                    }
                                    else
                                    {
                                        matrixData.Add(new MatrixRecord(cn, cp, stHead, m, vt, vi, type));
                                    }
                                }
                            }
                        }
                    }

                    foreach (var km in fileMonths)
                    {
                        var d = km.Value;
                        summaryData.Add(new SummaryRecord(km.Key, stHead, "Customer", d.GrossTaxable, d.GrossInvoice, d.InternalTaxableCustomer, d.InternalInvoiceCustomer, d.IsFallback));
                        summaryData.Add(new SummaryRecord(km.Key, stHead, "Supplier", d.GrossTaxable, d.GrossInvoice, d.InternalTaxableSupplier, d.InternalInvoiceSupplier, d.IsFallback));
                    }
                }

                foreach (var md in matrixData)
                {
                    md.IsRelatedParty = relatedPANs.Contains(md.PAN);
                    if (string.IsNullOrWhiteSpace(md.Name) || md.Name == "-")
                    {
                        if (panToNameMap.ContainsKey(md.PAN)) md.Name = panToNameMap[md.PAN];
                        else md.Name = md.PAN == "UNREGISTERED" ? "Consumer / Unregistered Sales" : "Unknown Counterparty";
                    }
                }

                var uniqueMonths = summaryData.Select(s => s.Month).Concat(matrixData.Select(m => m.Month))
                    .Distinct()
                    .Where(m => !string.IsNullOrEmpty(m))
                    .OrderBy(m => DateTime.ParseExact(m, "MMM-yy", System.Globalization.CultureInfo.InvariantCulture))
                    .ToList();

                var uniqueStates = summaryData.Select(s => s.State).Distinct().OrderBy(s => s).ToList();

                double auditGross = summaryData.Where(s => s.Type == "Customer").Sum(s => s.GrossTaxable);
                double auditInternal = summaryData.Where(s => s.Type == "Customer").Sum(s => s.InternalTaxable);
                AnsiConsole.MarkupLine($"      [white]Audit Verification Run:[/] Gross Revenue: [cyan]INR {auditGross:#,##0.00}[/] | Balanced Net: [green]INR {(auditGross - auditInternal):#,##0.00}[/]");

                // ESCAPED BRACKETS FOR SPECTRE
                AnsiConsole.MarkupLine("[yellow][[3/3]][/] Generating analytical reporting arrays...");
                string outputName = $"CONSOLIDATED_{currentPan}_{currentName}.xlsx";
                string outputPath = Path.Combine(currentFolder, outputName);

                using var outWb = new XLWorkbook();

                var configurations = new[]
                {
                    new NetConfig("Tax. Value - Internal Sales", "Customer", true, new[] { "Gross Revenue - Taxable Value", "Internal Sales - Taxable Value", "Net Revenue - Taxable Value" }),
                    new NetConfig("Inv. Value - Internal Sales", "Customer", false, new[] { "Gross Revenue - Invoice Value", "Internal Sales - Invoice Value", "Net Revenue - Invoice Value" }),
                    new NetConfig("Tax. Value - Internal Purchases", "Supplier", true, new[] { "Gross Revenue - Taxable Value", "Internal Purchases - Taxable Value", "Net Revenue - Taxable Value" }),
                    new NetConfig("Inv. Value - Internal Purchases", "Supplier", false, new[] { "Gross Revenue - Invoice Value", "Internal Purchases - Invoice Value", "Net Revenue - Invoice Value" })
                };

                foreach (var cfg in configurations)
                {
                    var ws = outWb.Worksheets.Add(cfg.SheetName);
                    int rowTracker = 1;

                    foreach (var block in new[] { "Gross", "Internal", "Net" })
                    {
                        string headerLabel = block == "Gross" ? cfg.Labels[0] : block == "Internal" ? cfg.Labels[1] : cfg.Labels[2];
                        ws.Cell(rowTracker, 1).SetValue(headerLabel).Style.Font.SetBold(true);
                        ws.Cell(rowTracker + 1, 1).SetValue("Month").Style.Font.SetBold(true);

                        int colIdx = 2;
                        foreach (var st in uniqueStates)
                        {
                            ws.Cell(rowTracker + 1, colIdx++).SetValue(st).Style.Font.SetBold(true);
                        }
                        ws.Cell(rowTracker + 1, colIdx).SetValue("Total").Style.Font.SetBold(true);

                        int dataRow = rowTracker + 2;
                        foreach (var m in uniqueMonths)
                        {
                            ws.Cell(dataRow, 1).SetValue(m).Style.NumberFormat.Format = "@";
                            colIdx = 2;
                            bool rowContainsFallback = false;

                            foreach (var st in uniqueStates)
                            {
                                var matches = summaryData.Where(s => s.Month == m && s.State == st && s.Type == cfg.TypeTarget).ToList();
                                double cellValue = 0;

                                if (matches.Count > 0)
                                {
                                    if (block == "Gross") cellValue = matches.Sum(s => cfg.IsTaxable ? s.GrossTaxable : s.GrossInvoice);
                                    else if (block == "Internal") cellValue = matches.Sum(s => cfg.IsTaxable ? s.InternalTaxable : s.InternalInvoice);
                                    else cellValue = matches.Sum(s => cfg.IsTaxable ? (s.GrossTaxable - s.InternalTaxable) : (s.GrossInvoice - s.InternalInvoice));

                                    if (matches.Any(s => s.IsFallback)) rowContainsFallback = true;
                                }

                                var targetCell = ws.Cell(dataRow, colIdx++);
                                targetCell.SetValue(cellValue);

                                if (rowContainsFallback && block != "Internal" && cellValue > 0)
                                {
                                    targetCell.Style.Font.Italic = true;
                                    targetCell.Style.Fill.BackgroundColor = XLColor.FromHtml("#FFF2CC");
                                }
                            }

                            string lastColLetter = GetColLetter(colIdx - 1);
                            ws.Cell(dataRow, colIdx).FormulaA1 = $"=SUM(B{dataRow}:{lastColLetter}{dataRow})";
                            dataRow++;
                        }
                        rowTracker = dataRow + 2;
                    }
                    ws.Columns().AdjustToContents();
                    ws.RangeUsed().Style.NumberFormat.Format = "#,##0.00";
                    ws.Column(1).Style.NumberFormat.Format = "@";
                }

                var matrixConfigs = new[]
                {
                    new MatrixConfig("Detailed_Customer_Taxable", "Customer", "T"),
                    new MatrixConfig("Detailed_Customer_Invoice", "Customer", "I"),
                    new MatrixConfig("Detailed_Supplier_Taxable", "Supplier", "T"),
                    new MatrixConfig("Detailed_Supplier_Invoice", "Supplier", "I")
                };

                foreach (var mCfg in matrixConfigs)
                {
                    var ws = outWb.Worksheets.Add(mCfg.SheetName);
                    ws.Cell(1, 1).SetValue("Party / State").Style.Font.Bold = true;
                    ws.Cell(1, 2).SetValue("PAN").Style.Font.Bold = true;

                    int colIdx = 3;
                    foreach (var m in uniqueMonths)
                    {
                        ws.Cell(1, colIdx++).SetValue(m).Style.Font.Bold = true;
                    }
                    ws.Cell(1, colIdx).SetValue("Total").Style.Font.Bold = true;

                    int dataRow = 2;
                    var groupedByPan = matrixData.Where(m => m.Type == mCfg.TypeTarget).GroupBy(m => m.PAN).ToList();

                    foreach (var panGroup in groupedByPan)
                    {
                        string firstPartyName = panGroup.First().Name;
                        bool isRp = panGroup.First().IsRelatedParty;

                        ws.Cell(dataRow, 1).SetValue(firstPartyName).Style.Font.Bold = true;
                        ws.Cell(dataRow, 2).SetValue(panGroup.Key).Style.Font.SetBold(true);
                        ws.Row(dataRow).Style.Fill.BackgroundColor = isRp ? XLColor.FromHtml("#FFF2CC") : XLColor.FromHtml("#E1E1E1");

                        colIdx = 3;
                        foreach (var m in uniqueMonths)
                        {
                            double v = panGroup.Where(g => g.Month == m).Sum(g => mCfg.ValTarget == "T" ? g.Taxable : g.Invoice);
                            if (v > 0) ws.Cell(dataRow, colIdx).SetValue(v);
                            colIdx++;
                        }
                        ws.Cell(dataRow, colIdx).FormulaA1 = $"=SUM(C{dataRow}:{GetColLetter(colIdx - 1)}{dataRow})";
                        
                        int parentRow = dataRow;
                        dataRow++;

                        var groupedByState = panGroup.GroupBy(g => g.State).ToList();
                        foreach (var stateGroup in groupedByState)
                        {
                            ws.Cell(dataRow, 1).SetValue($"   >> {stateGroup.Key}");
                            colIdx = 3;
                            foreach (var m in uniqueMonths)
                            {
                                double v = stateGroup.Where(g => g.Month == m).Sum(g => mCfg.ValTarget == "T" ? g.Taxable : g.Invoice);
                                if (v > 0) ws.Cell(dataRow, colIdx).SetValue(v);
                                colIdx++;
                            }
                            ws.Cell(dataRow, colIdx).FormulaA1 = $"=SUM(C{dataRow}:{GetColLetter(colIdx - 1)}{dataRow})";
                            dataRow++;
                        }
                        ws.Rows(parentRow + 1, dataRow - 1).Group();
                    }
                    ws.Columns().AdjustToContents();
                    ws.RangeUsed().Style.NumberFormat.Format = "#,##0.00";
                    ws.Column(1).Style.NumberFormat.Format = "@";
                    ws.Column(2).Style.NumberFormat.Format = "@";
                }

                var wsGlossary = outWb.Worksheets.Add("Audit_Glossary");
                wsGlossary.Cell("A1").SetValue("Reporting Ledger Color Key").Style.Font.Bold = true;
                wsGlossary.Cell("A3").Style.Fill.BackgroundColor = XLColor.FromHtml("#FFF2CC");
                wsGlossary.Cell("B3").SetValue("Related Party Configuration / Subledger Identifiers");
                wsGlossary.Cell("A4").Style.Fill.BackgroundColor = XLColor.FromHtml("#E1E1E1");
                wsGlossary.Cell("B4").SetValue("Third-Party Verified Operational Vectors");
                wsGlossary.Cell("A5").Style.Fill.BackgroundColor = XLColor.FromHtml("#FFF2CC");
                wsGlossary.Cell("A5").Style.Font.Italic = true;
                wsGlossary.Cell("B5").SetValue("GSTR1 Operational Fallback Values (Triggered when explicit GSTR3B data is filed as missing or 0)");
                wsGlossary.Columns().AdjustToContents();

                try
                {
                    outWb.SaveAs(outputPath);
                    // ESCAPED BRACKETS FOR SPECTRE
                    AnsiConsole.MarkupLine($"[green][[SUCCESS]][/] Structured ledger exported safely to: [yellow]{Markup.Escape(outputPath)}[/]\n");
                }
                catch (Exception fileEx)
                {
                    AnsiConsole.MarkupLine($"[red]!! Structural write execution blocked. Ensure file is not locked by Excel: {Markup.Escape(fileEx.Message)}[/]");
                }
            }

            AnsiConsole.MarkupLine("[green]Execution operations cleared. Pipeline context finalized.[/]");
            Console.WriteLine("\nPress any key to close...");
            Console.ReadKey();
        }

        private static double SafeDouble(XLCellValue value)
        {
            if (value.IsBlank) return 0;
            if (value.IsNumber) return value.GetNumber();
            string txt = value.ToString().Trim();
            if (double.TryParse(txt, out double res)) return res;
            return 0;
        }

        private static string GetColLetter(int n)
        {
            string s = "";
            while (n > 0)
            {
                int m = (n - 1) % 26;
                s = (char)(65 + m) + s;
                n = (n - m) / 26;
            }
            return s;
        }
    }

    // --- Core Architecture Domain Models ---
    public class FileMetadata(string file, string pan, string name, string gstin, string state, string suffix)
    {
        public string FilePath { get; set; } = file;
        public string PAN { get; set; } = pan;
        public string TradeName { get; set; } = name;
        public string GSTIN { get; set; } = gstin;
        public string StateCode { get; set; } = state;
        public string Suffix { get; set; } = suffix;
    }

    public class MonthData(double gt, double gi, bool fb)
    {
        public double GrossTaxable { get; set; } = gt;
        public double GrossInvoice { get; set; } = gi;
        public double InternalTaxableCustomer { get; set; } = 0;
        public double InternalInvoiceCustomer { get; set; } = 0;
        public double InternalTaxableSupplier { get; set; } = 0;
        public double InternalInvoiceSupplier { get; set; } = 0;
        public bool IsFallback { get; set; } = fb;
    }

    public class SummaryRecord(string m, string st, string t, double gt, double gi, double it, double ii, bool fb)
    {
        public string Month { get; set; } = m;
        public string State { get; set; } = st;
        public string Type { get; set; } = t;
        public double GrossTaxable { get; set; } = gt;
        public double GrossInvoice { get; set; } = gi;
        public double InternalTaxable { get; set; } = it;
        public double InternalInvoice { get; set; } = ii;
        public bool IsFallback { get; set; } = fb;
    }

    public class MatrixRecord(string n, string p, string st, string m, double t, double i, string ty)
    {
        public string Name { get; set; } = n;
        public string PAN { get; set; } = p;
        public string State { get; set; } = st;
        public string Month { get; set; } = m;
        public double Taxable { get; set; } = t;
        public double Invoice { get; set; } = i;
        public bool IsRelatedParty { get; set; } = false;
        public string Type { get; set; } = ty;
    }

    public class NetConfig(string sName, string target, bool isTax, string[] lbls)
    {
        public string SheetName { get; set; } = sName;
        public string TypeTarget { get; set; } = target;
        public bool IsTaxable { get; set; } = isTax;
        public string[] Labels { get; set; } = lbls;
    }

    public class MatrixConfig(string sName, string target, string valType)
    {
        public string SheetName { get; set; } = sName;
        public string TypeTarget { get; set; } = target;
        public string ValTarget { get; set; } = valType;
    }
}
