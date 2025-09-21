using System;
using System.IO;
using System.Data;
using System.Text;
using System.Globalization;
using System.Diagnostics;
using System.Collections.Generic;
using ClosedXML.Excel;
using ExcelDataReader;

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
                string exeFolder = Path.GetDirectoryName(Process.GetCurrentProcess().MainModule!.FileName!)!;
                string controlPath = Path.Combine(exeFolder, "control.xlsx");
                if (!File.Exists(controlPath))
                {
                    Console.WriteLine($"Файл control.xlsx не найден в папке {exeFolder}");
                    return;
                }
                Console.WriteLine("Файл control.xlsx найден и успешно прочитан.\n");

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
                    int lr = ws.LastRowUsed().RowNumber();
                    // -3: первые две строки конфигурации и одна строка условных обозначений
                    if (lr > 3) totalRows += (lr - 3);
                }
                int processed = 0;
                int sheetIndex = 0;

                foreach (var wsControl in sheets)
                {
                    sheetIndex++;

                    // Чтение настроек листа
                    string sheetDefaultFolder = wsControl.Cell("B1").GetString().Trim();
                    int decimals = 6;
                    if (int.TryParse(wsControl.Cell("B2").GetString().Trim(), out int d) && d >= 0)
                        decimals = d;

                    // Создаём лист в результате с таким же именем
                    var resultSheet = resultWorkbook.AddWorksheet(wsControl.Name);

                    int lastRow = wsControl.LastRowUsed().RowNumber();
                    int lastCol = wsControl.LastColumnUsed().ColumnNumber();

                    // Копируем структуру начиная с 3-й строки (условные обозначения)
                    for (int r = 3; r <= lastRow; r++)
                        for (int c = 1; c <= lastCol; c++)
                            resultSheet.Cell(r - 2, c).Value = wsControl.Cell(r, c).Value;

                    // --- Основная обработка начинается с 4-й строки control.xlsx ---
                    for (int r = 4; r <= lastRow; r++)
                    {
                        string folderPath   = wsControl.Cell(r, 1).GetString().Trim();
                        string fileName     = wsControl.Cell(r, 2).GetString().Trim();
                        string sheetName    = wsControl.Cell(r, 3).GetString().Trim();
                        
                        if (string.IsNullOrWhiteSpace(fileName) || string.IsNullOrWhiteSpace(sheetName))
                            continue;

                        if (!IsValidFileName(fileName))
                            continue; // пропустить строку

                        string targetFilePath = null;

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
                            if (files.Length > 0)
                                targetFilePath = files[0];
                        }

                        if (targetFilePath == null)
                        {
                            // пропускаем без вывода в консоль
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

                            // Если нашли через defaultFolder, пишем фактический путь в первую колонку
                            if (string.IsNullOrWhiteSpace(folderPath))
                                resultSheet.Cell(r - 2, 1).Value = targetFilePath;

                            for (int c = 4; c <= lastCol; c++)
                            {
                                string cellAddr = wsControl.Cell(r, c).GetString().Trim();
                                if (string.IsNullOrWhiteSpace(cellAddr)) continue;

                                object rawObj = GetCellValueWithFallback(table, cellAddr, targetFilePath, sheetName);
                                var resultCell = resultSheet.Cell(r - 2, c);

                                if (rawObj == null) continue;

                                if (rawObj is IConvertible conv &&
                                    (rawObj.GetType().IsPrimitive || rawObj is decimal))
                                {
                                    double num = conv.ToDouble(CultureInfo.InvariantCulture);
                                    resultCell.Value = num;
                                    resultCell.Style.NumberFormat.Format = decimals > 0
                                        ? "0." + new string('#', decimals)
                                        : "0";
                                }
                                else if (rawObj is DateTime dt)
                                {
                                    resultCell.Value = dt;
                                    resultCell.Style.DateFormat.Format = "yyyy-MM-dd HH:mm";
                                }
                                else
                                {
                                    // строка с возможным числом/датой
                                    string s = rawObj.ToString();
                                    if (double.TryParse(s, NumberStyles.Any, CultureInfo.CurrentCulture, out double n) ||
                                        double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out n))
                                    {
                                        resultCell.Value = n;
                                        resultCell.Style.NumberFormat.Format = decimals > 0
                                            ? "0." + new string('#', decimals)
                                            : "0";
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
                            // пропускаем ошибочный файл без сообщений
                        }
                        processed++;
                        ShowProgress(processed, totalRows);
                    }
                }

                Console.WriteLine(); // новая строка после прогресс-бара

                string timestamp = DateTime.Now.ToString("yyyy.MM.dd_HH-mm");
                string resultPath = Path.Combine(exeFolder, $"{timestamp} EA result.xlsx");
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
        }

        static bool IsValidPath(string? path)
        {
            if (string.IsNullOrWhiteSpace(path)) return false;
            // недопустимые символы в имени файла/папки
            char[] invalid = Path.GetInvalidPathChars();
            return path.IndexOfAny(invalid) < 0;
        }

        static bool IsValidFileName(string? name)
        {
            if (string.IsNullOrWhiteSpace(name)) return false;
            char[] invalid = Path.GetInvalidFileNameChars();
            return name.IndexOfAny(invalid) < 0;
        }

        // Вспомогательные методы
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
            reader = (ext == ".xls" || ext == ".xlsb")
                ? ExcelReaderFactory.CreateBinaryReader(stream)
                : ExcelReaderFactory.CreateReader(stream);

            var conf = new ExcelDataSetConfiguration
            {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration { UseHeaderRow = false }
            };
            using (reader) return reader.AsDataSet(conf);
        }

        static object GetCellValueWithFallback(DataTable table, string cellAddr, string file, string sheet)
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
                            // Получаем вычисленное значение ка object
                            object cv = cell.GetValue<object>();
                            string cvStr = cv?.ToString() ?? string.Empty;
                            if (!cvStr.TrimStart().StartsWith("="))
                                return cv;
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
            foreach (char c in letters.ToUpperInvariant())
            {
                sum = sum * 26 + (c - 'A' + 1);
            }
            return sum - 1;
        }
    }
}
