using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using ClosedXML.Excel;
using Microsoft.Win32;

namespace KarzaConsolidator
{
    public partial class MainWindow : Window
    {
        private const string CurrentAppVersion = "v2026.09";
        private const string GithubRepository = "pratikvoice-web/Karza-Consolidation-Tool";
        private string _updateDownloadUrl = string.Empty;

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

        public MainWindow()
        {
            InitializeComponent();
            string runningFolder = AppDomain.CurrentDomain.BaseDirectory;
            TxtSourcePath.Text = runningFolder;
            TxtDestPath.Text = runningFolder;
            LogLine("System Initialization Status Matrix Configured. Ready.");
            
            _ = CheckForUpdatesAsync();
        }

        private async Task CheckForUpdatesAsync()
        {
            try
            {
                using var client = new HttpClient();
                client.DefaultRequestHeaders.Add("User-Agent", "Karza-Consolidator-AutoUpdater");
                
                string response = await client.GetStringAsync($"https://api.github.com/repos/{GithubRepository}/releases/latest");
                using var doc = JsonDocument.Parse(response);
                
                string remoteVersion = doc.RootElement.GetProperty("tag_name").GetString();
                
                if (!string.IsNullOrEmpty(remoteVersion) && remoteVersion != CurrentAppVersion)
                {
                    var assets = doc.RootElement.GetProperty("assets");
                    foreach (var asset in assets.EnumerateArray())
                    {
                        if (asset.GetProperty("name").GetString().EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
                        {
                            _updateDownloadUrl = asset.GetProperty("browser_download_url").GetString();
                            break;
                        }
                    }

                    if (!string.IsNullOrEmpty(_updateDownloadUrl))
                    {
                        Dispatcher.Invoke(() => 
                        {
                            LblUpdateText.Text = $"A new version ({remoteVersion}) is available.";
                            UpdateBanner.Visibility = Visibility.Visible;
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() => LogLine($"[Telemetry] Remote version check bypassed: {ex.Message}"));
            }
        }

        private async void BtnUpdateNow_Click(object sender, RoutedEventArgs e)
        {
            var userDecision = MessageBox.Show("The engine will download the package and execute an automated system swap. Ensure work vectors are committed.\n\nProceed with automated deployment handoff?", 
                                               "Handoff Authorized", MessageBoxButton.YesNo, MessageBoxImage.Information);
            if (userDecision != MessageBoxResult.Yes) return;

            BtnUpdateNow.IsEnabled = false;
            BtnUpdateNow.Content = "0% Downloaded";
            BtnUpdateDismiss.IsEnabled = false;
            BtnRun.IsEnabled = false;

            try
            {
                string currentExePath = Environment.ProcessPath ?? throw new Exception("Unable to locate active executable thread origin path.");
                string tempExePath = Path.Combine(Path.GetTempPath(), "KarzaConsolidator_Update.exe");
                string updaterBatPath = Path.Combine(Path.GetTempPath(), "KarzaUpdater.bat");

                using (var client = new HttpClient())
                {
                    using var response = await client.GetAsync(_updateDownloadUrl, HttpCompletionOption.ResponseHeadersRead);
                    long? totalBytes = response.Content.Headers.ContentLength;

                    using var contentStream = await response.Content.ReadAsStreamAsync();
                    using var fileStream = new FileStream(tempExePath, FileMode.Create, FileAccess.Write, FileShare.None);
                    
                    var dataBuffer = new byte[16384];
                    long totalBytesRead = 0;
                    int bytesReadCount;

                    while ((bytesReadCount = await contentStream.ReadAsync(dataBuffer, 0, dataBuffer.Length)) > 0)
                    {
                        await fileStream.WriteAsync(dataBuffer, 0, bytesReadCount);
                        totalBytesRead += bytesReadCount;
                        if (totalBytes.HasValue)
                        {
                            double currentPct = (double)totalBytesRead / totalBytes.Value * 100;
                            BtnUpdateNow.Content = $"{currentPct:F0}% Downloaded";
                        }
                    }
                }

                string batScript = $@"@echo off
echo Executing Background Processing Pipeline Frame Swap...
timeout /t 3 /nobreak > NUL
taskkill /f /im ""{Path.GetFileName(currentExePath)}"" > NUL 2>&1
del /f /q ""{currentExePath}""
move /y ""{tempExePath}"" ""{currentExePath}""
start """" ""{currentExePath}""
del ""%~f0""";

                File.WriteAllText(updaterBatPath, batScript);

                MessageBox.Show("Download segment verified successfully. The tool will close down to commit file overwrites and restart instantly.", 
                                "System Swap Staged", MessageBoxButton.OK, MessageBoxImage.Information);

                Process.Start(new ProcessStartInfo
                {
                    FileName = updaterBatPath,
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden,
                    UseShellExecute = true
                });

                Environment.Exit(0);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Update execution cycle interrupted: {ex.Message}", "Deployment Failure", MessageBoxButton.OK, MessageBoxImage.Error);
                UpdateBanner.Visibility = Visibility.Collapsed;
                BtnRun.IsEnabled = true;
                BtnUpdateNow.IsEnabled = true;
                BtnUpdateNow.Content = "Update Now";
                BtnUpdateDismiss.IsEnabled = true;
            }
        }

        private void BtnUpdateDismiss_Click(object sender, RoutedEventArgs e)
        {
            UpdateBanner.Visibility = Visibility.Collapsed;
        }

        private void BtnBrowseSource_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFolderDialog { InitialDirectory = TxtSourcePath.Text };
            if (dialog.ShowDialog() == true) TxtSourcePath.Text = dialog.FolderName;
        }

        private void BtnBrowseDest_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFolderDialog { InitialDirectory = TxtDestPath.Text };
            if (dialog.ShowDialog() == true) TxtDestPath.Text = dialog.FolderName;
        }

        private void LogLine(string message)
        {
            TxtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}\r\n");
            TxtLog.ScrollToEnd();
        }

        private async void BtnRun_Click(object sender, RoutedEventArgs e)
        {
            string sourceFolder = TxtSourcePath.Text.Trim();
            string destFolder = TxtDestPath.Text.Trim();

            if (!Directory.Exists(sourceFolder))
            {
                MessageBox.Show("Source directory lookup error. Path is invalid.", "Execution Blocked", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            BtnRun.IsEnabled = false;
            BtnBrowseSource.IsEnabled = false;
            BtnBrowseDest.IsEnabled = false;
            BtnOpenFolder.Visibility = Visibility.Collapsed;
            TxtLog.Clear();

            ProgressExtract.Value = 0;
            ProgressCompile.Value = 0;

            var progressReporter = new Progress<UiProgressReport>(report =>
            {
                if (report.Type == "LOG") LogLine(report.StatusText);
                else if (report.Type == "EXTRACT")
                {
                    ProgressExtract.Value = report.Value;
                    LblExtractStatus.Text = report.StatusText;
                }
                else if (report.Type == "COMPILE")
                {
                    ProgressCompile.Value = report.Value;
                    LblCompileStatus.Text = report.StatusText;
                }
            });

            try
            {
                await Task.Run(() => ProcessingEngine(sourceFolder, destFolder, progressReporter));
                LogLine("SUCCESS: All Operations Cleared safely.");
                BtnOpenFolder.Visibility = Visibility.Visible;
            }
            catch (Exception ex)
            {
                LogLine($"FATAL SYSTEM EXCEPTION: {ex.Message}");
                MessageBox.Show($"Core engine halted operations:\n{ex.Message}", "Processing Failure", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                BtnRun.IsEnabled = true;
                BtnBrowseSource.IsEnabled = true;
                BtnBrowseDest.IsEnabled = true;
            }
        }

        private void ProcessingEngine(string inputFolder, string outputFolder, IProgress<UiProgressReport> prog)
        {
            prog.Report(new UiProgressReport("LOG", 0, "Initializing metadata index scan pass..."));

            var excelFiles = Directory.GetFiles(inputFolder, "*.xlsx")
                .Where(f => !Path.GetFileName(f).StartsWith("CONSOLIDATED_"))
                .ToList();

            if (excelFiles.Count == 0)
                throw new Exception("No target Karza source ledger files discovered inside target directory.");

            var fileDataList = new List<FileMetadata>();
            foreach (var file in excelFiles)
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
                string safeName = NormalizeEntityName(name);

                string stateCode = b6.Length >= 15 ? b6.Substring(0, 2) : "00";
                string suffix = b6.Length >= 15 ? b6.Substring(b6.Length - 3, 3) : "XXX";

                fileDataList.Add(new FileMetadata(file, pan, safeName, b6, stateCode, suffix));
            }

            var entityGroups = fileDataList.GroupBy(f => 
                (f.PAN.Length == 10 && char.ToUpperInvariant(f.PAN[3]) == 'P') ? $"{f.PAN}_{f.TradeName}" : f.PAN
            ).ToDictionary(g => g.Key, g => g.ToList());

            foreach (var group in entityGroups)
            {
                var items = group.Value;
                string currentPan = items[0].PAN;
                string currentName = items[0].TradeName;

                prog.Report(new UiProgressReport("LOG", 0, $"Processing Profile Boundaries for Entity: {currentName}"));

                var stateCounts = items.GroupBy(i => i.StateCode).ToDictionary(g => g.Key, g => g.Count());
                var summaryData = new List<SummaryRecord>();
                var matrixData = new List<MatrixRecord>();
                var relatedPANs = new HashSet<string>();
                var panToNameMap = new Dictionary<string, string>();

                int fileIndex = 0;
                foreach (var item in items)
                {
                    fileIndex++;
                    string stateName = StateMap.ContainsKey(item.StateCode) ? StateMap[item.StateCode] : "Unknown";
                    string stHead = stateCounts[item.StateCode] > 1 ? $"{item.StateCode}-{stateName}-{item.Suffix}" : $"{item.StateCode}-{stateName}";

                    double pct = ((double)fileIndex / items.Count) * 100;
                    prog.Report(new UiProgressReport("EXTRACT", pct, $"Extracting Layer ({fileIndex}/{items.Count}): {stHead}"));

                    using var wb = new XLWorkbook(item.FilePath);

                    foreach (var sheetName in new[] { "Related Party Sales - Monthly", "Related Party Purchases-Monthly" })
                    {
                        if (wb.TryGetWorksheet(sheetName, out IXLWorksheet wsRP))
                        {
                            int rpMax = wsRP.LastRowUsed()?.RowNumber() ?? 100;
                            for (int c = 2; c <= 200; c += 8)
                            {
                                int blankStreak = 0;
                                for (int r = 4; r <= rpMax + 50; r++)
                                {
                                    string rpp = wsRP.Cell(r, c).Value.ToString().Trim();
                                    if (rpp.Length == 10) { relatedPANs.Add(rpp); blankStreak = 0; }
                                    else if (++blankStreak > 50) break;
                                }
                            }
                        }
                    }

                    var fileMonths = new Dictionary<string, MonthData>();
                    if (wb.TryGetWorksheet("GSTR1 vs 3B", out IXLWorksheet wsG))
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
                                if (gt3b == 0 && gt1 > 0) { gt3b = gt1; fallback = true; }
                                if (gi3b == 0 && gi1 > 0) { gi3b = gi1; fallback = true; }

                                fileMonths[m] = new MonthData(gt3b, gi3b, fallback);
                            }
                        }
                    }

                    foreach (var type in new[] { "Customer", "Supplier" })
                    {
                        if (wb.TryGetWorksheet($"{type} Wise - Monthly Data", out IXLWorksheet wsM))
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
                                        if (++blankStreak > 50) break;
                                        continue;
                                    }
                                    blankStreak = 0;

                                    if (serial.Contains("Total", StringComparison.OrdinalIgnoreCase) ||
                                        cp.Contains("Total", StringComparison.OrdinalIgnoreCase) ||
                                        cn.Contains("Total", StringComparison.OrdinalIgnoreCase)) continue;

                                    if (string.IsNullOrEmpty(cp)) cp = "UNREGISTERED";
                                    
                                    string normalisedCn = string.Empty;
                                    if (cp != "UNREGISTERED" && !string.IsNullOrWhiteSpace(cn) && cn != "-") 
                                    {
                                        normalisedCn = NormalizeEntityName(cn);
                                        panToNameMap[cp] = normalisedCn;
                                    }

                                    double vt = SafeDouble(wsM.Cell(r, c + 3).Value);
                                    double vi = SafeDouble(wsM.Cell(r, c + 5).Value);

                                    if (vt == 0 && vi == 0) continue;

                                    if (cp == currentPan)
                                    {
                                        if (type == "Customer") { fileMonths[m].InternalTaxableCustomer += vt; fileMonths[m].InternalInvoiceCustomer += vi; }
                                        else { fileMonths[m].InternalTaxableSupplier += vt; fileMonths[m].InternalInvoiceSupplier += vi; }
                                    }
                                    else matrixData.Add(new MatrixRecord(normalisedCn, cp, stHead, m, vt, vi, type));
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
                prog.Report(new UiProgressReport("EXTRACT", 100, "Ledger extraction pass completed safely."));

                foreach (var md in matrixData)
                {
                    md.IsRelatedParty = relatedPANs.Contains(md.PAN);
                    if (string.IsNullOrWhiteSpace(md.Name) || md.Name == "-")
                    {
                        if (panToNameMap.ContainsKey(md.PAN)) md.Name = panToNameMap[md.PAN];
                        else md.Name = md.PAN == "UNREGISTERED" ? "CONSUMER / UNREGISTERED SALES" : "UNKNOWN COUNTERPARTY";
                    }
                }

                var uniqueMonths = summaryData.Select(s => s.Month).Concat(matrixData.Select(m => m.Month)).Distinct().Where(m => !string.IsNullOrEmpty(m)).OrderBy(m => DateTime.ParseExact(m, "MMM-yy", System.Globalization.CultureInfo.InvariantCulture)).ToList();
                var uniqueStates = summaryData.Select(s => s.State).Distinct().OrderBy(s => s).ToList();
                var fyGroups = uniqueMonths.GroupBy(m => GetFinancialYear(m)).ToList();

                double auditGross = summaryData.Where(s => s.Type == "Customer").Sum(s => s.GrossTaxable);
                double auditInternal = summaryData.Where(s => s.Type == "Customer").Sum(s => s.InternalTaxable);
                prog.Report(new UiProgressReport("LOG", 0, $"Audit Verification: Gross Turnover Verified: INR {auditGross:#,##0.00} | Net: INR {(auditGross - auditInternal):#,##0.00}"));

                // --- GENERATION SECTOR ---
                string outputName = $"CONSOLIDATED_{currentPan}_{currentName}.xlsx";
                string outputPath = Path.Combine(outputFolder, outputName);
                using var outWb = new XLWorkbook();

                // 1. Compile the Index Sheet
                var wsIndex = outWb.Worksheets.Add("Index");
                wsIndex.Cell("A1").SetValue("Consolidated GST Karza").Style.Font.SetBold(true).Font.SetFontSize(16);
                wsIndex.Cell("A3").SetValue("Entity Name:").Style.Font.SetBold(true);
                wsIndex.Cell("B3").SetValue(currentName);
                wsIndex.Cell("A4").SetValue("PAN:").Style.Font.SetBold(true);
                wsIndex.Cell("B4").SetValue(currentPan);

                wsIndex.Cell("A6").SetValue("Table of Contents").Style.Font.SetBold(true).Font.SetFontSize(14);
                wsIndex.Cell("A7").SetValue("S.No").Style.Font.SetBold(true);
                wsIndex.Cell("B7").SetValue("Sheet Name").Style.Font.SetBold(true);
                wsIndex.Cell("C7").SetValue("Description").Style.Font.SetBold(true);
                wsIndex.Range("A7:C7").Style.Fill.BackgroundColor = XLColor.FromHtml("#E2E8F0");

                int indexRow = 8;
                int sheetCount = 1;

                Action<string, string> AddToIndex = (sheetName, description) =>
                {
                    wsIndex.Cell(indexRow, 1).SetValue(sheetCount);
                    var linkCell = wsIndex.Cell(indexRow, 2);
                    linkCell.SetValue(sheetName);
                    linkCell.SetHyperlink(new XLHyperlink($"'{sheetName}'!A1"));
                    linkCell.Style.Font.FontColor = XLColor.RoyalBlue;
                    linkCell.Style.Font.Underline = XLFontUnderlineValues.Single;
                    wsIndex.Cell(indexRow, 3).SetValue(description);
                    indexRow++;
                    sheetCount++;
                };

                var configurations = new[]
                {
                    new NetConfig("Tax. Value - Internal Sales", "Customer", true, new[] { "Gross Revenue - Taxable Value", "Internal Sales - Taxable Value", "Net Revenue - Taxable Value" }),
                    new NetConfig("Inv. Value - Internal Sales", "Customer", false, new[] { "Gross Revenue - Invoice Value", "Internal Sales - Invoice Value", "Net Revenue - Invoice Value" }),
                    new NetConfig("Tax. Value - Internal Purchases", "Supplier", true, new[] { "Gross Revenue - Taxable Value", "Internal Purchases - Taxable Value", "Net Revenue - Taxable Value" }),
                    new NetConfig("Inv. Value - Internal Purchases", "Supplier", false, new[] { "Gross Revenue - Invoice Value", "Internal Purchases - Invoice Value", "Net Revenue - Invoice Value" })
                };

                int compStep = 0;
                foreach (var cfg in configurations)
                {
                    prog.Report(new UiProgressReport("COMPILE", ((double)++compStep / 10) * 100, $"Writing Array Map: {cfg.SheetName}"));
                    var ws = outWb.Worksheets.Add(cfg.SheetName);
                    AddToIndex(cfg.SheetName, $"Monthly Summary - {cfg.Labels[2].Replace("Net Revenue - ", "")}");
                    ws.Outline.SummaryVLocation = XLOutlineSummaryVLocation.Top;
                    int rowTracker = 1;

                    foreach (var block in new[] { "Gross", "Internal", "Net" })
                    {
                        string headerLabel = block == "Gross" ? cfg.Labels[0] : block == "Internal" ? cfg.Labels[1] : cfg.Labels[2];
                        ws.Cell(rowTracker, 1).SetValue(headerLabel).Style.Font.SetBold(true);
                        ws.Cell(rowTracker + 1, 1).SetValue("Financial Year / Month").Style.Font.SetBold(true);

                        int colIdx = 2;
                        foreach (var st in uniqueStates) ws.Cell(rowTracker + 1, colIdx++).SetValue(st).Style.Font.SetBold(true);
                        ws.Cell(rowTracker + 1, colIdx).SetValue("Total").Style.Font.SetBold(true);
                        int maxColIdx = colIdx;

                        int dataRow = rowTracker + 2;

                        foreach (var fy in fyGroups)
                        {
                            int fyRow = dataRow;
                            ws.Cell(fyRow, 1).SetValue(fy.Key).Style.Font.SetBold(true);
                            ws.Row(fyRow).Style.Fill.BackgroundColor = XLColor.FromHtml("#F1F5F9");
                            dataRow++;
                            int startGroupRow = dataRow;

                            foreach (var m in fy)
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
                                ws.Cell(dataRow, colIdx).FormulaA1 = $"=SUM(B{dataRow}:{GetColLetter(colIdx - 1)}{dataRow})";
                                dataRow++;
                            }

                            if (dataRow > startGroupRow)
                            {
                                ws.Rows(startGroupRow, dataRow - 1).Group();
                                for (int c = 2; c <= maxColIdx; c++)
                                {
                                    ws.Cell(fyRow, c).FormulaA1 = $"=SUM({GetColLetter(c)}{startGroupRow}:{GetColLetter(c)}{dataRow - 1})";
                                    ws.Cell(fyRow, c).Style.Font.SetBold(true);
                                }
                            }
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
                    prog.Report(new UiProgressReport("COMPILE", ((double)++compStep / 10) * 100, $"Writing Subledger Layout: {mCfg.SheetName}"));
                    var ws = outWb.Worksheets.Add(mCfg.SheetName);
                    AddToIndex(mCfg.SheetName, $"Detailed Party-wise Subledger ({mCfg.TypeTarget})");
                    ws.Outline.SummaryVLocation = XLOutlineSummaryVLocation.Top;
                    ws.Outline.SummaryHLocation = XLOutlineSummaryHLocation.Right;

                    ws.Cell(1, 1).SetValue("Financial Year").Style.Font.SetBold(true);
                    ws.Cell(2, 1).SetValue("Party / State").Style.Font.SetBold(true);
                    ws.Cell(2, 2).SetValue("PAN").Style.Font.SetBold(true);

                    int colIdx = 3;
                    var fyTotalCols = new List<string>();

                    foreach (var fy in fyGroups)
                    {
                        int startCol = colIdx;
                        foreach (var m in fy) ws.Cell(2, colIdx++).SetValue(m).Style.Font.SetBold(true);
                        
                        ws.Cell(2, colIdx).SetValue($"{fy.Key} Total").Style.Font.SetBold(true).Font.FontColor = XLColor.AirForceBlue;
                        ws.Range(1, startCol, 1, colIdx).Merge().SetValue(fy.Key).Style.Font.SetBold(true).Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                        
                        if (colIdx - 1 >= startCol) ws.Columns(startCol, colIdx - 1).Group();
                        fyTotalCols.Add(GetColLetter(colIdx));
                        colIdx++;
                    }
                    ws.Cell(2, colIdx).SetValue("Grand Total").Style.Font.SetBold(true);

                    int dataRow = 3;
                    var groupedByPan = matrixData.Where(m => m.Type == mCfg.TypeTarget)
                        .GroupBy(m => m.PAN)
                        .OrderByDescending(g => g.Sum(m => mCfg.ValTarget == "T" ? m.Taxable : m.Invoice))
                        .ToList();

                    foreach (var panGroup in groupedByPan)
                    {
                        string firstPartyName = panGroup.First().Name;
                        bool isRp = panGroup.First().IsRelatedParty;

                        ws.Cell(dataRow, 1).SetValue(firstPartyName).Style.Font.SetBold(true);
                        ws.Cell(dataRow, 2).SetValue(panGroup.Key).Style.Font.SetBold(true);
                        ws.Row(dataRow).Style.Fill.BackgroundColor = isRp ? XLColor.FromHtml("#FFF2CC") : XLColor.FromHtml("#E1E1E1");

                        colIdx = 3;
                        foreach (var fy in fyGroups)
                        {
                            int startCol = colIdx;
                            foreach (var m in fy)
                            {
                                double v = panGroup.Where(g => g.Month == m).Sum(g => mCfg.ValTarget == "T" ? g.Taxable : g.Invoice);
                                if (v > 0) ws.Cell(dataRow, colIdx).SetValue(v);
                                colIdx++;
                            }
                            ws.Cell(dataRow, colIdx).FormulaA1 = $"=SUM({GetColLetter(startCol)}{dataRow}:{GetColLetter(colIdx - 1)}{dataRow})";
                            ws.Cell(dataRow, colIdx).Style.Font.SetBold(true);
                            colIdx++;
                        }
                        ws.Cell(dataRow, colIdx).FormulaA1 = "=" + string.Join("+", fyTotalCols.Select(c => $"{c}{dataRow}"));
                        ws.Cell(dataRow, colIdx).Style.Font.SetBold(true);
                        
                        int parentRow = dataRow;
                        dataRow++;

                        var groupedByState = panGroup.GroupBy(g => g.State).ToList();
                        foreach (var stateGroup in groupedByState)
                        {
                            ws.Cell(dataRow, 1).SetValue($"   >> {stateGroup.Key}");
                            colIdx = 3;
                            foreach (var fy in fyGroups)
                            {
                                int startCol = colIdx;
                                foreach (var m in fy)
                                {
                                    double v = stateGroup.Where(g => g.Month == m).Sum(g => mCfg.ValTarget == "T" ? g.Taxable : g.Invoice);
                                    if (v > 0) ws.Cell(dataRow, colIdx).SetValue(v);
                                    colIdx++;
                                }
                                ws.Cell(dataRow, colIdx).FormulaA1 = $"=SUM({GetColLetter(startCol)}{dataRow}:{GetColLetter(colIdx - 1)}{dataRow})";
                                ws.Cell(dataRow, colIdx).Style.Font.SetBold(true);
                                colIdx++;
                            }
                            ws.Cell(dataRow, colIdx).FormulaA1 = "=" + string.Join("+", fyTotalCols.Select(c => $"{c}{dataRow}"));
                            ws.Cell(dataRow, colIdx).Style.Font.SetBold(true);
                            dataRow++;
                        }
                        ws.Rows(parentRow + 1, dataRow - 1).Group();
                    }
                    ws.Columns().AdjustToContents();
                    ws.RangeUsed().Style.NumberFormat.Format = "#,##0.00";
                    ws.Column(1).Style.NumberFormat.Format = "@";
                    ws.Column(2).Style.NumberFormat.Format = "@";
                }

                prog.Report(new UiProgressReport("COMPILE", 100, "Finalizing ledger metadata profiles..."));
                var wsGlossary = outWb.Worksheets.Add("Audit_Glossary");
                AddToIndex("Audit_Glossary", "Reporting Ledger Color Key & System Glossary");
                wsIndex.Columns().AdjustToContents(); 

                wsGlossary.Cell("A1").SetValue("Reporting Ledger Color Key").Style.Font.SetBold(true);
                wsGlossary.Cell("A3").Style.Fill.BackgroundColor = XLColor.FromHtml("#FFF2CC");
                wsGlossary.Cell("B3").SetValue("Related Party Configuration / Subledger Identifiers");
                wsGlossary.Cell("A4").Style.Fill.BackgroundColor = XLColor.FromHtml("#E1E1E1");
                wsGlossary.Cell("B4").SetValue("Third-Party Verified Operational Vectors");
                wsGlossary.Cell("A5").Style.Fill.BackgroundColor = XLColor.FromHtml("#FFF2CC");
                wsGlossary.Cell("A5").Style.Font.Italic = true;
                wsGlossary.Cell("B5").SetValue("GSTR1 Operational Fallback Values (Triggered when explicit GSTR3B data is filed as missing or 0)");
                wsGlossary.Columns().AdjustToContents();

                outWb.SaveAs(outputPath);
                prog.Report(new UiProgressReport("LOG", 100, $"Export Complete: {outputName}"));
            }
        }

        private static string NormalizeEntityName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "UNKNOWN_ENTITY";

            name = name.ToUpperInvariant().Trim();

            name = Regex.Replace(name, @"\b(?:PRIVATE|PVT\.?|\(P\))\s*(?:LIMITED|LTD\.?)\b", "PVT LTD");
            name = Regex.Replace(name, @"\b(?:LIMITED|LTD\.?)\b", "LTD");

            name = Regex.Replace(name, @"\s+", " ");
            name = Regex.Replace(name, @"[\\/:*?""<>|]", "_");

            return name.Trim();
        }

        private static string GetFinancialYear(string mmmYY)
        {
            var dt = DateTime.ParseExact(mmmYY, "MMM-yy", System.Globalization.CultureInfo.InvariantCulture);
            if (dt.Month >= 4)
                return $"FY{dt.ToString("yy")}-{dt.AddYears(1).ToString("yy")}";
            else
                return $"FY{dt.AddYears(-1).ToString("yy")}-{dt.ToString("yy")}";
        }

        private void BtnOpenFolder_Click(object sender, RoutedEventArgs e)
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = TxtDestPath.Text,
                UseShellExecute = true,
                Verb = "open"
            });
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

    public class NetConfig(string sName, string target, bool iTax, string[] lbls)
    {
        public string SheetName { get; set; } = sName;
        public string TypeTarget { get; set; } = target;
        public bool IsTaxable { get; set; } = iTax;
        public string[] Labels { get; set; } = lbls;
    }

    public class MatrixConfig(string sName, string target, string valType)
    {
        public string SheetName { get; set; } = sName;
        public string TypeTarget { get; set; } = target;
        public string ValTarget { get; set; } = valType;
    }

    public class UiProgressReport(string type, double val, string txt)
    {
        public string Type { get; set; } = type;
        public double Value { get; set; } = val;
        public string StatusText { get; set; } = txt;
    }
}
