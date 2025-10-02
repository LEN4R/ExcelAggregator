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
        // Кэш открытых книг
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
                // Папка, где лежит exe
                string originalExePath = GetOriginalExePath();
                string originalExeFolder = Path.GetDirectoryName(originalExePath)!;

                // --- Поиск control_*.xlsx ---
                string? controlPath = Directory.GetFiles(originalExeFolder, "control_*.xlsx").FirstOrDefault();

                if (controlPath == null)
                {
                    Console.WriteLine($"Файл control_*.xlsx не найден рядом с программой: {originalExeFolder}");
                    Console.Write("Укажите папку, где находится control_*.xlsx или Enter для выхода: ");
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

                // Итоговая книга
                var resultWorkbook = new XLWorkbook();

                using var controlWorkbook = new XLWorkbook(controlPath);

                // --- Сбор всех листов ---
                var sheets = controlWorkbook.Worksheets;
                int totalRows = 0;
                foreach (var ws in sheets)
                {
                    var lastRowUsed = ws.LastRowUsed();
                    if (lastRowUsed == null) continue;
                    int lr = lastRowUsed.RowNumber();
                    if (lr > 3) totalRows += (lr - 3);
                }

                int processed = 0;

                foreach (var wsControl in sheets)
                {
                    // Чтение настроек листа
                    string sheetDefaultFolder = wsControl.Cell("B1").GetString().Trim();
                    int decimals = 6;
                    if (int.TryParse(wsControl.Cell("B2").GetString().Trim(), out int d) && d >= 0)
                        decimals = d;

                    // Создаём лист в результате с таким же именем
                    var resultSheet = resultWorkbook.AddWorksheet(wsControl.Name);

                    int lastRow = wsControl.LastRowUsed().RowNumber();
                    int lastCol = wsControl.LastColumnUsed().ColumnNumber();

                    // --- Копируем структуру начиная с 3-й строки ---
                    for (int r = 3; r <= lastRow; r++)
                    {
                        for (int c = 1; c <= lastCol; c++)
                        {
                            var srcCell = wsControl.Cell(r, c);
                            var dstCell = resultSheet.Cell(r - 2, c);

                            if (srcCell.HasFormula)
                                dstCell.FormulaA1 = srcCell.FormulaA1; // копируем формулу
                            else
                                dstCell.Value = srcCell.Value; // копируем значение
                        }
                    }

                    // Закрепляем первые 3 строки
                    resultSheet.SheetView.FreezeRows(3);

                    // --- Основная обработка с 4-й строки ---
                    for (int r = 4; r <= lastRow; r++)
                    {
                        string folderPath = wsControl.Cell(r, 1).GetString().Trim();
                        string fileName = wsControl.Cell(r, 2).GetString().Trim();
                        string sheetName = wsControl.Cell(r, 3).GetString().Trim();

                        if (string.IsNullOrWhiteSpace(fileName) || string.IsNullOrWhiteSpace(sheetName)) continue;
                        if (!IsValidFileName(fileName)) continue;

                        string? targetFilePath = null;

                        // 1) путь из колонки A
                        if (!string.IsNullOrWhiteSpace(folderPath) && IsValidPath(folderPath))
                        {
                            string path = Path.Combine(folderPath, fileName);
                            if (File.Exists(path)) targetFilePath = path;
                        }
                        // 2) поиск в папке по умолчанию B1 (рекурсивно)
                        else if (!string.IsNullOrWhiteSpace(sheetDefaultFolder))
                        {
                            var files = Directory.GetFiles(sheetDefaultFolder, fileName, SearchOption.AllDirectories);
                            if (files.Length > 0) targetFilePath = files[0];
                        }

                        if (targetFilePath == null)
                        {
                            processed++;
                            ShowProgress(processed, totalRows);
                            continue;
                        }

                        try
                        {
                            var ds = ReadExcelFile(targetFilePath);
                            if (!ds.Tables.Contains(sheetName))
                            {
                                processed++;
                                ShowProgress(processed, totalRows);
                                continue;
                            }

                            var table = ds.Tables[sheetName];

                            if (string.IsNullOrWhiteSpace(folderPath))
                                resultSheet.Cell(r - 2, 1).Value = targetFilePath;

                            for (int c = 4; c <= lastCol; c++)
                            {
                                string cellAddr = wsControl.Cell(r, c).GetString().Trim();
                                if (string.IsNullOrWhiteSpace(cellAddr)) continue;

                                object? rawObj = GetCellValueWithFallback(table, cellAddr, targetFilePath, sheetName);
                                var resultCell = resultSheet.Cell(r - 2, c);

                                // Не затираем формулы из control-файла
                                if (resultCell.HasFormula) continue;

                                if (rawObj == null) continue;

                                if (rawObj is IConvertible conv && (rawObj.GetType().IsPrimitive || rawObj is decimal))
                                {
                                    double num = conv.ToDouble(CultureInfo.InvariantCulture);
                                    resultCell.Value = num;
                                    resultCell.Style.NumberFormat.Format = decimals > 0 ? "0." + new string('#', decimals) : "0";
                                }
                                else if (rawObj is DateTime dt)
                                {
                                    resultCell.Value = dt;
                                    resultCell.Style.DateFormat.Format = "yyyy-MM-dd HH:mm";
                                }
                                else
                                {
                                    string s = rawObj.ToString()!;
                                    if (double.TryParse(s, NumberStyles.Any, CultureInfo.CurrentCulture, out double n) ||
                                        double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out n))
                                    {
                                        resultCell.Value = n;
                                        resultCell.Style.NumberFormat.Format = decimals > 0 ? "0." + new string('#', decimals) : "0";
                                    }
                                    else if (DateTime.TryParse(s, CultureInfo.CurrentCulture, DateTimeStyles.None, out DateTime d1) ||
                                             DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.None, out d1))
                                    {
                                        resultCell.Value = d1;
                                        resultCell.Style.DateFormat.Format = "yyyy-MM-dd HH:mm";
                                    }
                                    else
                                    {
                                        resultCell.Value = s;
                                    }
                                }
                            }
                        }
                        catch
                        {
                            // пропускаем ошибочный файл
                        }
                        processed++;
                        ShowProgress(processed, totalRows);
                    }

                    // Авторасширение столбцов
                    resultSheet.Columns().AdjustToContents();
                }

                Console.WriteLine(); // новая строка после прогресс-бара

                string timestamp = DateTime.Now.ToString("yyyy.MM.dd_HH-mm");
                string resultPath = Path.Combine(originalExeFolder, $"{timestamp} EA result.xlsx");
                resultWorkbook.SaveAs(resultPath);

                Console.WriteLine($"\nПроцесс успешно выполнен.\nРезультат сохранён в:\n{resultPath}");
                Console.WriteLine("\n(с) Галиев Ленар\nИсходный код: https://github.com/LEN4R/ExcelAggregator/");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\nПроцесс завершился с ошибкой: {ex.Message}");
            }
            finally
            {
                CleanupCache();
            }

            Console.WriteLine("Нажмите Enter для выхода...");
            Console.ReadLine();
        }

        // НОВЫЙ МЕТОД: Получение пути к оригинальному EXE файлу
        static string GetOriginalExePath()
        {
            try
            {
                string processPath = Process.GetCurrentProcess().MainModule!.FileName!;
                if (!processPath.Contains(@"\Temp\", StringComparison.OrdinalIgnoreCase))
                    return processPath;
            }
            catch { }

            string baseDir = AppContext.BaseDirectory;
            if (!baseDir.Contains(@"\Temp\", StringComparison.OrdinalIgnoreCase))
            {
                var exeFiles = Directory.GetFiles(baseDir, "*.exe");
                if (exeFiles.Length > 0) return exeFiles[0];
                return Path.Combine(baseDir, Process.GetCurrentProcess().ProcessName + ".exe");
            }

            try
            {
                string assemblyPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
                if (!string.IsNullOrEmpty(assemblyPath) && !assemblyPath.Contains(@"\Temp\", StringComparison.OrdinalIgnoreCase))
                    return assemblyPath;
            }
            catch { }

            return Process.GetCurrentProcess().MainModule!.FileName!;
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
            IExcelDataReader reader;
            string ext = Path.GetExtension(path).ToLowerInvariant();
            reader = (ext == ".xls" || ext == ".xlsb") ? ExcelReaderFactory.CreateBinaryReader(stream) : ExcelReaderFactory.CreateReader(stream);
            var conf = new ExcelDataSetConfiguration { ConfigureDataTable = _ => new ExcelDataTableConfiguration { UseHeaderRow = false } };
            using (reader) return reader.AsDataSet(conf);
        }

        static object? GetCellValueWithFallback(DataTable table, string cellAddr, string file, string sheet)
        {
            try
            {
                string letters = "", numbers = "";
                foreach (char ch in cellAddr)
                {
                    if (char.IsLetter(ch)) letters += ch;
                    else if (char.IsDigit(ch)) numbers += ch;
                }
                if (!int.TryParse(numbers, out int row) || string.IsNullOrEmpty(letters)) return null;

                int col = ColumnLetterToNumber(letters);
                if (row - 1 >= table.Rows.Count || col >= table.Columns.Count) return null;

                object val = table.Rows[row - 1][col];
                if (val is string s && s.TrimStart().StartsWith("="))
                {
                    string ext = Path.GetExtension(file).ToLowerInvariant();
                    if (ext == ".xlsx" || ext == ".xlsm")
                    {
                        if (!_workbookCache.TryGetValue(file, out var wb))
                        {
                            wb = new XLWorkbook(file);
                            _workbookCache[file] = wb;
                        }
                        var ws = wb.Worksheet(sheet);
                        var cell = ws.Cell(cellAddr);
                        if (!cell.IsEmpty())
                        {
                            object cv = cell.GetValue<object>();
                            string cvStr = cv?.ToString() ?? string.Empty;
                            if (!cvStr.TrimStart().StartsWith("=")) return cv;
                        }
                    }
                }
                return val;
            }
            catch { return null; }
        }

        static int ColumnLetterToNumber(string letters)
        {
            int sum = 0;
            foreach (char c in letters.ToUpperInvariant()) sum = sum * 26 + (c - 'A' + 1);
            return sum - 1;
        }
    }
}
