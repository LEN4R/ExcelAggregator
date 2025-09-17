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
        // Кэш для открытых книг
        private static readonly Dictionary<string, XLWorkbook> _workbookCache = new Dictionary<string, XLWorkbook>();
        
        static void Main(string[] args)
        {
            // Регистрация кодировок (нужно для чтения старых xls)
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            Console.OutputEncoding = Encoding.UTF8;

            // Обработка Ctrl+C
            Console.CancelKeyPress += (sender, e) => 
            {
                Console.WriteLine("\nПрервано пользователем");
                CleanupCache();
                Environment.Exit(0);
            };

            try
            {
                // 1) Папка по умолчанию (ввод)
                Console.WriteLine("Введите путь к папке по умолчанию:");
                string defaultFolder = Console.ReadLine()?.Trim() ?? "";
                if (!Directory.Exists(defaultFolder))
                {
                    Console.WriteLine("Папка не существует!");
                    return;
                }

                // 2) Путь к exe (настоящему исполняемому файлу, работает с single-file)
                string exePath = Process.GetCurrentProcess().MainModule.FileName;
                string exeFolder = Path.GetDirectoryName(exePath) ?? AppContext.BaseDirectory;

                // 3) Сколько знаков после запятой
                Console.Write("Введите количество знаков после запятой (по умолчанию 6): ");
                string input = Console.ReadLine();
                int decimals = 6;
                if (!string.IsNullOrWhiteSpace(input))
                {
                    if (!int.TryParse(input, out decimals) || decimals < 0)
                    {
                        Console.WriteLine("Некорректное значение, использую 6.");
                        decimals = 6;
                    }
                }

                // 4) control.xlsx рядом с exe
                string controlPath = Path.Combine(exeFolder, "control.xlsx");
                if (!File.Exists(controlPath))
                {
                    Console.WriteLine($"Файл control.xlsx не найден в папке {exeFolder}");
                    return;
                }

                // 5) Создаём рабочую книгу результата и копируем структуру control.xlsx
                var resultWorkbook = new XLWorkbook();
                var resultSheet = resultWorkbook.AddWorksheet("Результат");

                using (var controlWorkbook = new XLWorkbook(controlPath))
                {
                    var wsControl = controlWorkbook.Worksheet(1);

                    int lastRow = wsControl.LastRowUsed().RowNumber();
                    int lastCol = wsControl.LastColumnUsed().ColumnNumber();

                    // Копируем структуру (значения заголовков и адреса ячеек)
                    for (int r = 1; r <= lastRow; r++)
                        for (int c = 1; c <= lastCol; c++)
                            resultSheet.Cell(r, c).Value = wsControl.Cell(r, c).Value;

                    // Обрабатываем строки (начиная со 2-й)
                    for (int r = 2; r <= lastRow; r++)
                    {
                        string folderPath = wsControl.Cell(r, 1).GetString();
                        string fileName = wsControl.Cell(r, 2).GetString();
                        string sheetName = wsControl.Cell(r, 3).GetString();

                        if (string.IsNullOrWhiteSpace(fileName) || string.IsNullOrWhiteSpace(sheetName))
                            continue;

                        string searchFolder = string.IsNullOrWhiteSpace(folderPath) ? defaultFolder : folderPath;
                        string targetFilePath = null;

                        if (string.IsNullOrWhiteSpace(folderPath))
                        {
                            // ищем во всех подпапках defaultFolder
                            var files = Directory.GetFiles(defaultFolder, fileName, SearchOption.AllDirectories);
                            if (files.Length > 0)
                                targetFilePath = files[0];
                        }
                        else
                        {
                            // ищем только в указанной папке (без подпапок)
                            string path = Path.Combine(searchFolder, fileName);
                            if (File.Exists(path))
                                targetFilePath = path;
                        }

                        if (targetFilePath == null)
                        {
                            Console.WriteLine($"Файл {fileName} не найден.");
                            continue;
                        }

                        try
                        {
                            // Читаем файл через ExcelDataReader (все форматы)
                            DataSet ds = ReadExcelFile(targetFilePath);
                            if (!ds.Tables.Contains(sheetName))
                            {
                                Console.WriteLine($"Лист '{sheetName}' не найден в файле {fileName}");
                                continue;
                            }

                            var table = ds.Tables[sheetName];

                            for (int c = 4; c <= lastCol; c++)
                            {
                                string cellAddr = wsControl.Cell(r, c).GetString();
                                if (string.IsNullOrWhiteSpace(cellAddr))
                                    continue;

                                object rawObj = GetCellValueWithFallback(table, cellAddr, targetFilePath, sheetName);

                                if (rawObj == null)
                                {
                                    resultSheet.Cell(r, c).Value = "Ошибка";
                                    continue;
                                }

                                // Запись в result, сохраняя тип
                                var resultCell = resultSheet.Cell(r, c);

                                // Числа
                                if (rawObj is IConvertible convertible && 
                                    (rawObj.GetType().IsPrimitive || rawObj is decimal))
                                {
                                    double num = convertible.ToDouble(CultureInfo.InvariantCulture);
                                    resultCell.Value = num;
                                    // Применяем формат в зависимости от decimals
                                    if (decimals > 0)
                                        resultCell.Style.NumberFormat.Format = "0." + new string('#', decimals);
                                    else
                                        resultCell.Style.NumberFormat.Format = "0";
                                }
                                // Даты
                                else if (rawObj is DateTime dt)
                                {
                                    resultCell.Value = dt;
                                    resultCell.Style.DateFormat.Format = "yyyy-MM-dd HH:mm";
                                }
                                // Булевы
                                else if (rawObj is bool b)
                                {
                                    resultCell.Value = b;
                                }
                                // Строки (включая длинные числа, которые не хотим терять)
                                else
                                {
                                    // s — string, значение исходной ячейки
                                    string s = rawObj.ToString();

                                    // Попробуем сначала распознать число (с локалью пользователя, затем Invariant)
                                    double parsedNumber;
                                    bool isNumber = double.TryParse(s, NumberStyles.Any, CultureInfo.CurrentCulture, out parsedNumber);
                                    if (!isNumber)
                                    {
                                        isNumber = double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out parsedNumber);
                                    }

                                    if (isNumber)
                                    {
                                        var cell = resultCell;
                                        cell.Value = parsedNumber;
                                        if (decimals > 0)
                                            cell.Style.NumberFormat.Format = "0." + new string('#', decimals);
                                        else
                                            cell.Style.NumberFormat.Format = "0";
                                    }
                                    else
                                    {
                                        // Попробуем распознать дату (локаль пользователя, затем Invariant)
                                        DateTime parsedDate;
                                        bool isDate = DateTime.TryParse(s, CultureInfo.CurrentCulture, DateTimeStyles.None, out parsedDate);
                                        if (!isDate)
                                        {
                                            isDate = DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.None, out parsedDate);
                                        }

                                        if (isDate)
                                        {
                                            resultCell.Value = parsedDate;
                                            resultCell.Style.DateFormat.Format = "yyyy-MM-dd HH:mm";
                                        }
                                        else
                                        {
                                            // Оставляем текст
                                            resultCell.Value = s;
                                        }
                                    }

                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Ошибка при обработке файла {fileName}, лист {sheetName}: {ex.Message}");
                        }
                    }
                }

                // Сохраняем результат рядом с exe
                string timestamp = DateTime.Now.ToString("yyyy.MM.dd_HH-mm");
                string resultPath = Path.Combine(exeFolder, $"{timestamp} EA result.xlsx");
                resultWorkbook.SaveAs(resultPath);

                Console.WriteLine($"Готово! Результат сохранен в:\n{resultPath}");
                Console.WriteLine("\n(с) Галиев Ленар\nИсходный код: https://github.com/LEN4R/ExcelAggregator/");
            }
            finally
            {
                // Всегда очищаем кэш, даже при ошибках
                CleanupCache();
            }
        }

        /// <summary>
        /// Очистка кэша книг
        /// </summary>
        static void CleanupCache()
        {
            foreach (var wb in _workbookCache.Values)
            {
                wb.Dispose();
            }
            _workbookCache.Clear();
        }

        /// <summary>
        /// Читает Excel с помощью ExcelDataReader и возвращает DataSet (все листы).
        /// Открывает файл с FileShare.ReadWrite чтобы читать открытые файлы.
        /// </summary>
        static DataSet ReadExcelFile(string path)
        {
            using var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

            IExcelDataReader reader;
            string ext = Path.GetExtension(path).ToLowerInvariant();

            // ExcelDataReader умеет xls,xlsx,xlsm,xlsb (CreateReader понимает большинство)
            if (ext == ".xls")
                reader = ExcelReaderFactory.CreateBinaryReader(stream);
            else if (ext == ".xlsb")
                reader = ExcelReaderFactory.CreateBinaryReader(stream);
            else
                reader = ExcelReaderFactory.CreateReader(stream); // xlsx,xlsm

            var conf = new ExcelDataSetConfiguration
            {
                ConfigureDataTable = (_) => new ExcelDataTableConfiguration { UseHeaderRow = false }
            };

            using (reader)
            {
                return reader.AsDataSet(conf);
            }
        }

        /// <summary>
        /// Возвращает значение ячейки (object) из DataTable. Если значение похоже на формулу,
        /// пытается получить кешированное вычисленное значение через ClosedXML (только для xlsx/xlsm).
        /// Если ClosedXML не даёт результата — возвращает то, что есть в table.
        /// </summary>
        static object GetCellValueWithFallback(DataTable table, string cellAddr, string targetFilePath, string sheetName)
        {
            if (string.IsNullOrWhiteSpace(cellAddr))
                return null;

            try
            {
                // Разделяем буквы и цифры
                string colLetter = "";
                string rowNumber = "";
                foreach (char ch in cellAddr)
                {
                    if (char.IsLetter(ch)) colLetter += ch;
                    else if (char.IsDigit(ch)) rowNumber += ch;
                }

                if (string.IsNullOrWhiteSpace(colLetter) || string.IsNullOrWhiteSpace(rowNumber))
                    return null;

                int colIndex = ColumnLetterToNumber(colLetter);
                if (!int.TryParse(rowNumber, out int rowIndexNumber))
                    return null;
                int rowIndex = rowIndexNumber - 1;

                if (rowIndex < 0 || colIndex < 0 ||
                    rowIndex >= table.Rows.Count || colIndex >= table.Columns.Count)
                    return null;

                object val = table.Rows[rowIndex][colIndex];
                if (val == null || val == DBNull.Value)
                    return null;

                // Если в таблице строка и она начинается с '=' => возможно формула
                // Проверяем, похоже ли значение на формулу (строка, начинающаяся с '=')
                if (val is string sVal && !string.IsNullOrWhiteSpace(sVal) && sVal.TrimStart().StartsWith("="))
                {
                    string ext = Path.GetExtension(targetFilePath).ToLowerInvariant();

                    // Только для xlsx/xlsm используем ClosedXML для получения вычисленного значения
                    if (ext == ".xlsx" || ext == ".xlsm")
                    {
                        try
                        {
                            // Получаем или создаём кэшированную книгу
                            if (!_workbookCache.TryGetValue(targetFilePath, out XLWorkbook wb))
                            {
                                wb = new XLWorkbook(targetFilePath);
                                _workbookCache[targetFilePath] = wb;
                            }

                            var ws = wb.Worksheet(sheetName);
                            var cell = ws.Cell(cellAddr);

                            if (!cell.IsEmpty())
                            {
                                object cv = cell.Value;

                                // Если значение не формула — возвращаем
                                if (!(cv is string cvStr && cvStr.TrimStart().StartsWith("=")))
                                {
                                    return cv;
                                }
                            }
                        }
                        catch
                        {
                            // Игнорируем ошибки и вернём исходное значение
                        }
                    }
                }
                return val;
            }
            catch
            {
                return null;
            }
        }

        // Преобразуем A..Z.. в индекс (0-based)
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