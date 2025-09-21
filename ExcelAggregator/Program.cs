using ClosedXML.Excel;
using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
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
                string controlPath = Path.Combine(originalExeFolder, "control.xlsx");

                // Проверяем наличие control.xlsx рядом с exe
                if (!File.Exists(controlPath))
                {
                    Console.WriteLine($"Файл control.xlsx не найден рядом с программой: {originalExeFolder}");
                    Console.Write("Укажите папку, где находится control.xlsx или Enter для выхода: ");
                    string? userDir = Console.ReadLine();

                    if (string.IsNullOrWhiteSpace(userDir))
                    {
                        Console.WriteLine("Работа программы прекращена. Нажмите Enter для выхода.");
                        Console.ReadLine();
                        return;
                    }

                    string altPath = Path.Combine(userDir.Trim('"'), "control.xlsx");
                    if (!File.Exists(altPath))
                    {
                        Console.WriteLine("В указанной папке файл control.xlsx не найден. Завершение работы.");
                        Console.ReadLine();
                        return;
                    }

                    controlPath = altPath;
                }

                Console.WriteLine($"Файл control.xlsx найден: {controlPath}\n");

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
                    if (lr > 3) totalRows += (lr - 3); // -3: первые две строки конфигурации и строка условных обозначений
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

                    // Копируем структуру начиная с 3-й строки
                    for (int r = 3; r <= lastRow; r++)
                        for (int c = 1; c <= lastCol; c++)
                            resultSheet.Cell(r - 2, c).Value = wsControl.Cell(r, c).Value;

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
            // Способ 1: Через Process.MainModule (самый надежный для single-file)
            try
            {
                string processPath = Process.GetCurrentProcess().MainModule.FileName;

                // Проверяем, не временный ли это путь
                if (!processPath.Contains(@"\Temp\") && !processPath.Contains(@"\temp\"))
                {
                    return processPath;
                }
            }
            catch
            {
                // Если не получилось, пробуем другие способы
            }

            // Способ 2: Через AppContext.BaseDirectory (может вернуть временную папку)
            string baseDir = AppContext.BaseDirectory;
            if (!baseDir.Contains(@"\Temp\") && !baseDir.Contains(@"\temp\"))
            {
                // Ищем EXE файл в этой папке
                var exeFiles = Directory.GetFiles(baseDir, "*.exe");
                if (exeFiles.Length > 0)
                {
                    return exeFiles[0]; // возвращаем первый найденный EXE
                }
                return Path.Combine(baseDir, Process.GetCurrentProcess().ProcessName + ".exe");
            }
            
            // Способ 3: Через Assembly.Location (последний вариант)
            try
            {
                string assemblyPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
                if (!string.IsNullOrEmpty(assemblyPath) && !assemblyPath.Contains(@"\Temp\"))
                {
                    return assemblyPath;
                }
            }
            catch
            {
                // Если все способы не сработали
            }

            // Если ничего не помогло, возвращаем то, что есть
            return Process.GetCurrentProcess().MainModule.FileName;
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
