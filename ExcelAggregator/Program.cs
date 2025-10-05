using ClosedXML.Excel;
using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace ExcelAggregator
{
    class Program
    {
        private static readonly Dictionary<string, XLWorkbook> _workbookCache = new();

        static void Main()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            Console.OutputEncoding = Encoding.UTF8;

            Console.CancelKeyPress += (s, e) =>
            {
                CleanupCache();
                Environment.Exit(0);
            };

            try
            {
                string originalExePath = GetOriginalExePath();
                string originalExeFolder = Path.GetDirectoryName(originalExePath)!;

                // --- Поиск control_*.xlsx ---
                string? controlPath = Directory.GetFiles(originalExeFolder, "control_*.xlsx").FirstOrDefault();

                if (controlPath == null)
                {
                    Console.WriteLine($"Файл настройки не найден рядом с программой!");
                    Console.Write("Укажите папку, где находится настройки или Enter для выхода: ");
                    string? userDir = Console.ReadLine();

                    if (string.IsNullOrWhiteSpace(userDir))
                    {
                        Console.WriteLine("Работа программы прекращена. Нажмите Enter для выхода.");
                        Console.ReadLine();
                        return;
                    }

                    controlPath = Directory.GetFiles(userDir.Trim('"'), "control_*.xlsx").FirstOrDefault();
                    if (controlPath == null)
                    {
                        Console.WriteLine("В указанной папке файл control_*.xlsx не найден. Завершение работы.");
                        Console.ReadLine();
                        return;
                    }
                }

                Console.WriteLine($"Файл control найден: {controlPath}\n");

                var resultWorkbook = new XLWorkbook();
                using var controlWorkbook = new XLWorkbook(controlPath);

                int totalRows = 0;
                foreach (var ws in controlWorkbook.Worksheets)
                {
                    var lastRowUsed = ws.LastRowUsed();
                    if (lastRowUsed == null) continue;
                    int lr = lastRowUsed.RowNumber();
                    if (lr > 6) totalRows += (lr - 6);
                }

                int processed = 0;

                foreach (var wsControl in controlWorkbook.Worksheets)
                {
                    string defaultFolder = wsControl.Cell("B1").GetString().Trim();
                    int decimals = int.TryParse(wsControl.Cell("B2").GetString().Trim(), out int d) ? d : 6;
                    string autoFit = wsControl.Cell("B3").GetString().Trim().ToLower();
                    bool doAutoFit = autoFit == "да" || autoFit == "yes" || autoFit == "true";

                    var resultSheet = resultWorkbook.AddWorksheet(wsControl.Name);
                    int lastRow = wsControl.LastRowUsed().RowNumber();
                    int lastCol = wsControl.LastColumnUsed().ColumnNumber();

                    // --- Копирование структуры начиная с 6-й строки ---
                    for (int r = 6; r <= lastRow; r++)
                    {
                        for (int c = 1; c <= lastCol; c++)
                        {
                            var src = wsControl.Cell(r, c);
                            var dst = resultSheet.Cell(r - 5, c);

                            string fileName = wsControl.Cell(r, 2).GetString().Trim();
                            if (string.IsNullOrEmpty(fileName))
                            {
                                // если строка вспомогательная → сохраняем формулы
                                if (src.HasFormula)
                                    dst.FormulaA1 = src.FormulaA1;
                                else
                                    dst.Value = src.Value;
                            }
                            else
                            {
                                // строки с файлами → только значения
                                dst.Value = src.Value;
                            }
                        }
                    }

                    // --- Закрепляем только первую строку ---
                    resultSheet.SheetView.FreezeRows(1);

                    // --- Обработка строк с файлами ---
                    for (int r = 7; r <= lastRow; r++)
                    {
                        string folder = wsControl.Cell(r, 1).GetString().Trim();
                        string fileName = wsControl.Cell(r, 2).GetString().Trim();
                        string sheetName = wsControl.Cell(r, 3).GetString().Trim();

                        if (string.IsNullOrWhiteSpace(fileName) || string.IsNullOrWhiteSpace(sheetName)) continue;
                        if (!IsValidFileName(fileName)) continue;

                        string? path = null;
                        if (!string.IsNullOrWhiteSpace(folder) && IsValidPath(folder))
                        {
                            string full = Path.Combine(folder, fileName);
                            if (File.Exists(full)) path = full;
                        }
                        else if (!string.IsNullOrWhiteSpace(defaultFolder))
                        {
                            var found = Directory.GetFiles(defaultFolder, fileName, SearchOption.AllDirectories);
                            if (found.Length > 0) path = found[0];
                        }

                        if (path == null)
                        {
                            processed++;
                            ShowProgress(processed, totalRows);
                            continue;
                        }

                        try
                        {
                            var ds = ReadExcelFile(path);
                            if (!ds.Tables.Contains(sheetName))
                            {
                                processed++;
                                ShowProgress(processed, totalRows);
                                continue;
                            }

                            var table = ds.Tables[sheetName];

                            if (string.IsNullOrWhiteSpace(folder))
                                resultSheet.Cell(r - 5, 1).Value = path;

                            for (int c = 4; c <= lastCol; c++)
                            {
                                string cellAddr = wsControl.Cell(r, c).GetString().Trim();
                                if (string.IsNullOrWhiteSpace(cellAddr)) continue;

                                object? val = GetCellValueWithFallback(table, cellAddr, path, sheetName);
                                var dest = resultSheet.Cell(r - 5, c);

                                if (dest.HasFormula) continue;
                                if (val == null) continue;

                                if (val is IConvertible conv && (val.GetType().IsPrimitive || val is decimal))
                                {
                                    double num = conv.ToDouble(CultureInfo.InvariantCulture);
                                    dest.Value = num;
                                    dest.Style.NumberFormat.Format = decimals > 0 ? "0." + new string('#', decimals) : "0";
                                }
                                else if (val is DateTime dt)
                                {
                                    dest.Value = dt;
                                    dest.Style.DateFormat.Format = "yyyy-MM-dd HH:mm";
                                }
                                else
                                {
                                    string s = val.ToString()!;
                                    if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out double n))
                                    {
                                        dest.Value = n;
                                        dest.Style.NumberFormat.Format = decimals > 0 ? "0." + new string('#', decimals) : "0";
                                    }
                                    else if (DateTime.TryParse(s, out DateTime dd))
                                    {
                                        dest.Value = dd;
                                        dest.Style.DateFormat.Format = "yyyy-MM-dd HH:mm";
                                    }
                                    else
                                    {
                                        dest.Value = s;
                                    }
                                }
                            }
                        }
                        catch { }

                        processed++;
                        ShowProgress(processed, totalRows);
                    }

                    if (doAutoFit)
                    {
                        resultSheet.Columns().AdjustToContents();
                        resultSheet.Rows().AdjustToContents();
                    }
                }

                Console.WriteLine();

                string timestamp = DateTime.Now.ToString("yyyy.MM.dd_HH-mm");
                string resultPath = Path.Combine(originalExeFolder, $"{timestamp} EA result.xlsx");
                resultWorkbook.SaveAs(resultPath);

                Console.WriteLine($"\nПроцесс завершён.\nФайл сохранён: {resultPath}");
                Console.WriteLine("\n(с) Галиев Ленар");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\nОшибка: {ex.Message}");
            }
            finally
            {
                CleanupCache();
            }

            Console.WriteLine("Нажмите Enter для выхода...");
            Console.ReadLine();
        }

        static string GetOriginalExePath()
        {
            try
            {
                string p = Process.GetCurrentProcess().MainModule!.FileName!;
                if (!p.Contains(@"\Temp\", StringComparison.OrdinalIgnoreCase))
                    return p;
            }
            catch { }

            string baseDir = AppContext.BaseDirectory;
            var exe = Directory.GetFiles(baseDir, "*.exe").FirstOrDefault();
            return exe ?? Path.Combine(baseDir, Process.GetCurrentProcess().ProcessName + ".exe");
        }

        static bool IsValidPath(string? path) => !string.IsNullOrWhiteSpace(path) && path.IndexOfAny(Path.GetInvalidPathChars()) < 0;
        static bool IsValidFileName(string? name) => !string.IsNullOrWhiteSpace(name) && name.IndexOfAny(Path.GetInvalidFileNameChars()) < 0;

        static void ShowProgress(int done, int total)
        {
            double percent = total == 0 ? 100 : (double)done / total * 100;
            int width = 40;
            int filled = (int)(percent / 100 * width);
            string bar = new string('#', filled).PadRight(width, '-');
            Console.CursorLeft = 0;
            Console.Write($"[{bar}] {done}/{total}  {percent,5:0.0}%");
        }

        static void CleanupCache()
        {
            foreach (var wb in _workbookCache.Values) wb.Dispose();
            _workbookCache.Clear();
        }

        static DataSet ReadExcelFile(string path)
        {
            using var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            IExcelDataReader reader = Path.GetExtension(path).ToLowerInvariant() switch
            {
                ".xls" or ".xlsb" => ExcelReaderFactory.CreateBinaryReader(stream),
                _ => ExcelReaderFactory.CreateReader(stream)
            };
            var conf = new ExcelDataSetConfiguration { ConfigureDataTable = _ => new ExcelDataTableConfiguration { UseHeaderRow = false } };
            using (reader) return reader.AsDataSet(conf);
        }

        static object? GetCellValueWithFallback(DataTable table, string addr, string file, string sheet)
        {
            try
            {
                string letters = new(addr.Where(char.IsLetter).ToArray());
                string numbers = new(addr.Where(char.IsDigit).ToArray());
                if (!int.TryParse(numbers, out int row) || string.IsNullOrEmpty(letters)) return null;

                int col = ColumnLetterToNumber(letters);
                if (row - 1 >= table.Rows.Count || col >= table.Columns.Count) return null;

                object val = table.Rows[row - 1][col];
                if (val is string s && s.TrimStart().StartsWith("="))
                {
                    if (!_workbookCache.TryGetValue(file, out var wb))
                    {
                        wb = new XLWorkbook(file);
                        _workbookCache[file] = wb;
                    }
                    var ws = wb.Worksheet(sheet);
                    var cell = ws.Cell(addr);
                    if (!cell.IsEmpty()) return cell.Value;
                }
                return val;
            }
            catch { return null; }
        }

        static int ColumnLetterToNumber(string letters)
        {
            int sum = 0;
            foreach (char c in letters.ToUpperInvariant())
                sum = sum * 26 + (c - 'A' + 1);
            return sum - 1;
        }
    }
}
