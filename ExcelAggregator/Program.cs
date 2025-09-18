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
        // Кэш для открытых книг ClosedXML
        private static readonly Dictionary<string, XLWorkbook> _workbookCache = new();

        static void Main(string[] args)
        {
            // Регистрация кодировок для ExcelDataReader
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            Console.OutputEncoding = Encoding.UTF8;

            // Обработка Ctrl+C
            Console.CancelKeyPress += (s, e) =>
            {
                Console.WriteLine("\nПрервано пользователем");
                CleanupCache();
                Environment.Exit(0);
            };

            try
            {
                // Папка исполняемого файла
                string exePath = Process.GetCurrentProcess().MainModule!.FileName!;
                string exeFolder = Path.GetDirectoryName(exePath) ?? AppContext.BaseDirectory;

                // ---- Чтение control.xlsx ----
                string controlPath = Path.Combine(exeFolder, "control.xlsx");
                if (!File.Exists(controlPath))
                {
                    Console.WriteLine($"Файл control.xlsx не найден в папке {exeFolder}");
                    return;
                }
                Console.WriteLine($"Файл control.xlsx найден и успешно прочитан.\n");

                // Открываем control.xlsx
                using var controlWorkbook = new XLWorkbook(controlPath);
                var wsControl = controlWorkbook.Worksheet(1);

                int lastRow = wsControl.LastRowUsed().RowNumber();
                int lastCol = wsControl.LastColumnUsed().ColumnNumber();

                // Чтение настроек: B1 – путь по умолчанию, B2 – количество знаков
                string defaultFolder = wsControl.Cell("B1").GetString().Trim();
                int decimals = 6;
                int.TryParse(wsControl.Cell("B2").GetString().Trim(), out decimals);
                if (decimals < 0) decimals = 6;

                // Создаём книгу для результата
                var resultWorkbook = new XLWorkbook();
                var resultSheet = resultWorkbook.AddWorksheet("Результат");

                // Копируем заголовки (все строки целиком, если хотите – можно копировать только третью)
                for (int r = 1; r <= lastRow; r++)
                    for (int c = 1; c <= lastCol; c++)
                        resultSheet.Cell(r, c).Value = wsControl.Cell(r, c).Value;

                // ---- Подготовка прогресса ----
                int totalRows = lastRow - 3;    // начинаем с 4-й строки
                int processed = 0;
                Console.WriteLine("Обработка файлов:\n");

                // Основной цикл (начиная с 4-й строки)
                for (int r = 4; r <= lastRow; r++)
                {
                    string folderPath = wsControl.Cell(r, 1).GetString().Trim();
                    string fileName = wsControl.Cell(r, 2).GetString().Trim();
                    string sheetName = wsControl.Cell(r, 3).GetString().Trim();

                    if (string.IsNullOrWhiteSpace(fileName) || string.IsNullOrWhiteSpace(sheetName))
                    {
                        processed++;
                        ShowProgress(processed, totalRows);
                        continue;
                    }

                    // Поиск файла
                    string? targetFilePath = null;

                    if (!string.IsNullOrWhiteSpace(folderPath))
                    {
                        string p = Path.Combine(folderPath, fileName);
                        if (File.Exists(p)) targetFilePath = p;
                    }
                    else if (!string.IsNullOrEmpty(defaultFolder))
                    {
                        var files = Directory.GetFiles(defaultFolder, fileName, SearchOption.AllDirectories);
                        if (files.Length > 0) targetFilePath = files[0];
                    }

                    if (targetFilePath == null)
                    {
                        Console.WriteLine($"\nФайл {fileName} не найден.");
                        processed++;
                        ShowProgress(processed, totalRows);
                        continue;
                    }

                    try
                    {
                        // Читаем Excel
                        DataSet ds = ReadExcelFile(targetFilePath);
                        if (!ds.Tables.Contains(sheetName))
                        {
                            Console.WriteLine($"\nЛист '{sheetName}' не найден в {fileName}");
                            processed++;
                            ShowProgress(processed, totalRows);
                            continue;
                        }

                        var table = ds.Tables[sheetName];

                        for (int c = 4; c <= lastCol; c++)
                        {
                            string cellAddr = wsControl.Cell(r, c).GetString().Trim();
                            if (string.IsNullOrEmpty(cellAddr))
                                continue;

                            object? rawObj = GetCellValueWithFallback(table, cellAddr, targetFilePath, sheetName);
                            var resultCell = resultSheet.Cell(r, c);

                            if (rawObj == null)
                            {
                                resultCell.Value = "Ошибка";
                                continue;
                            }

                            // Запись с форматами
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
                                string s = rawObj.ToString() ?? "";
                                if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out double parsedNum))
                                {
                                    resultCell.Value = parsedNum;
                                    resultCell.Style.NumberFormat.Format = decimals > 0
                                        ? "0." + new string('#', decimals)
                                        : "0";
                                }
                                else if (DateTime.TryParse(s, out DateTime parsedDate))
                                {
                                    resultCell.Value = parsedDate;
                                    resultCell.Style.DateFormat.Format = "yyyy-MM-dd HH:mm";
                                }
                                else
                                {
                                    resultCell.Value = s;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"\nОшибка при обработке {fileName}: {ex.Message}");
                    }

                    processed++;
                    ShowProgress(processed, totalRows);
                }

                Console.WriteLine(); // перевод строки после прогресс-бара

                // Сохраняем результат
                string timestamp = DateTime.Now.ToString("yyyy.MM.dd_HH-mm");
                string resultPath = Path.Combine(exeFolder, $"{timestamp} EA result.xlsx");
                resultWorkbook.SaveAs(resultPath);

                Console.WriteLine("\nПроцесс успешно выполнен!");
                Console.WriteLine($"Результат сохранён в:\n{resultPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\nПроцесс завершился с ошибкой: {ex.Message}");
            }
            finally
            {
                CleanupCache();
                Console.WriteLine("\n(с) Галиев Ленар\nИсходный код: https://github.com/LEN4R/ExcelAggregator/");
            }
        }

        // === Вспомогательные методы ===

        static void ShowProgress(int done, int total)
        {
            double percent = total > 0 ? (double)done / total * 100 : 100;
            int width = 40;
            int filled = (int)(percent / 100 * width);
            string bar = new string('#', filled).PadRight(width, '-');
            Console.CursorLeft = 0;
            Console.Write($"[{bar}] {percent,6:0.0}%");
        }

        static void CleanupCache()
        {
            foreach (var wb in _workbookCache.Values)
                wb.Dispose();
            _workbookCache.Clear();
        }

        static DataSet ReadExcelFile(string path)
        {
            using var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            IExcelDataReader reader;
            string ext = Path.GetExtension(path).ToLowerInvariant();

            reader = ext switch
            {
                ".xls" or ".xlsb" => ExcelReaderFactory.CreateBinaryReader(stream),
                _ => ExcelReaderFactory.CreateReader(stream),
            };

            var conf = new ExcelDataSetConfiguration
            {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration { UseHeaderRow = false }
            };
            using (reader) return reader.AsDataSet(conf);
        }

        static object? GetCellValueWithFallback(DataTable table, string cellAddr,
                                               string targetFilePath, string sheetName)
        {
            if (string.IsNullOrWhiteSpace(cellAddr)) return null;
            try
            {
                string colLetter = "", rowNumber = "";
                foreach (char ch in cellAddr)
                {
                    if (char.IsLetter(ch)) colLetter += ch;
                    else if (char.IsDigit(ch)) rowNumber += ch;
                }

                if (!int.TryParse(rowNumber, out int rowIndexNumber)) return null;
                int colIndex = ColumnLetterToNumber(colLetter);
                int rowIndex = rowIndexNumber - 1;
                if (rowIndex < 0 || colIndex < 0 ||
                    rowIndex >= table.Rows.Count || colIndex >= table.Columns.Count)
                    return null;

                object val = table.Rows[rowIndex][colIndex];
                if (val == null || val == DBNull.Value) return null;

                if (val is string sVal && sVal.TrimStart().StartsWith("="))
                {
                    string ext = Path.GetExtension(targetFilePath).ToLowerInvariant();
                    if (ext == ".xlsx" || ext == ".xlsm")
                    {
                        try
                        {
                            if (!_workbookCache.TryGetValue(targetFilePath, out var wb))
                            {
                                wb = new XLWorkbook(targetFilePath);
                                _workbookCache[targetFilePath] = wb;
                            }
                            var ws = wb.Worksheet(sheetName);
                            var cell = ws.Cell(cellAddr);
                            if (!cell.IsEmpty())
                            {
                                object cv = cell.Value;
                                if (!(cv is string cvStr && cvStr.TrimStart().StartsWith("=")))
                                    return cv;
                            }
                        }
                        catch { /* игнор */ }
                    }
                }
                return val;
            }
            catch { return null; }
        }

        static int ColumnLetterToNumber(string colLetter)
        {
            int sum = 0;
            foreach (char c in colLetter.ToUpperInvariant())
            {
                sum *= 26;
                sum += (c - 'A' + 1);
            }
            return sum - 1;
        }
    }
}
