using System;
using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;
using ExcelRyan.Model;
using Newtonsoft.Json;

namespace ExcelRyan
{
    internal class Program
    {
        private static List<InvoiceEntry> _allEntries = new List<InvoiceEntry>();
        private static Dictionary<string, AssesedClient> _clients = new Dictionary<string, AssesedClient>();
        private static Settings _settings;

        static Dictionary<string, IXLStyle> _schedule2ColumnStyles = new Dictionary<string, IXLStyle>();
        static Dictionary<string, IXLStyle> _schedule3ColumnStyles = new Dictionary<string, IXLStyle>();

        public static void Main(string[] args)
        {
            try
            {
                if (args.Length == 0)
                {
                    throw new ArgumentException("Please specify a settings file (you can use drag/drop)");
                }

                var settingsFileContent = File.ReadAllText(args[0]);
                _settings = JsonConvert.DeserializeObject<Settings>(settingsFileContent);

                Execute();
            }
            finally
            {
                Console.WriteLine("Press any key to close");
                Console.ReadKey();
            }
        }

        static void Execute()
        {
            _allEntries = LoadRawData();

            Console.WriteLine("Data loaded. Processing each clients found");

            foreach (var kvp in _clients)
            {
                CreatePackage(kvp.Key, kvp.Value);
            }
        }

        static List<InvoiceEntry> LoadRawData()
        {
            Console.WriteLine($"Loading workbook {_settings.InputDocumentPath}");
            Console.WriteLine("This might take a while ^^");

            using (var workbook = new XLWorkbook(_settings.InputDocumentPath))
            {
                Console.WriteLine($"{_settings.InputDocumentPath} loaded.");

                foreach (var sheet in workbook.Worksheets)
                {
                    Console.WriteLine($"Processing sheet {sheet.Name}");

                    Settings.WorksheetColumns sheetSettings;
                    if (!_settings.Worksheets.TryGetValue(sheet.Name, out sheetSettings))
                    {
                        Console.WriteLine($"{sheet.Name}, skipped, not found in settings file");
                        continue;
                    }

                    var rowCount = sheet.RowCount();
                    var currentRow = sheetSettings.FirstRow;
                    while (currentRow <= rowCount &&
                           !string.IsNullOrEmpty(sheet.Cell(currentRow, "A").GetString()))
                    {
                        Console.WriteLine($"Processing row {currentRow}");

                        var invoiceEntry = new InvoiceEntry
                        {
                            Date = sheet.Cell(currentRow, sheetSettings.Date).GetString(),
                            InvoiceId = sheet.Cell(currentRow, sheetSettings.InvoiceId).GetString(),
                            ClientId = sheet.Cell(currentRow, sheetSettings.ClientId).GetString(),
                            Amount = sheet.Cell(currentRow, sheetSettings.Amount).GetDouble(),
                            Assesed = sheet.Cell(currentRow, sheetSettings.Assessed).GetDouble(),
                        };

                        _allEntries.Add(invoiceEntry);

                        if (!_clients.ContainsKey(invoiceEntry.ClientId))
                        {
                            _clients[invoiceEntry.ClientId] = new AssesedClient();
                        }
                        _clients[invoiceEntry.ClientId].AddEntry(invoiceEntry);

                        currentRow++;
                    }

                    Console.WriteLine($"{sheet.Name} done. Processed {currentRow - sheetSettings.FirstRow + 1} rows.");
                }
            }
            return _allEntries;
        }

        static void CreatePackage(string clientID, AssesedClient client)
        {
            Console.WriteLine($"Creating package for {clientID}");

            var filePath = String.Format(Path.Combine(_settings.OutputFolder, $"{clientID}.xlsx"));

#if !DEBUG
            if (File.Exists(filePath))
            {
                throw new FileLoadException($"{filePath} already exists");
            }
#endif

            Directory.CreateDirectory(_settings.OutputFolder);
            File.Copy(_settings.TemplateDocumentPath, filePath, true);



            using (var workbook = new XLWorkbook(filePath))
            {
                CreateSchedule2(workbook, client);
                CreateSchedule3(workbook, client);

                workbook.Save(true);
            }
        }

        static void CreateSchedule2(XLWorkbook workbook, AssesedClient client)
        {
            IXLWorksheet schedule2Sheet;
            if (workbook.TryGetWorksheet("Schedule 2", out schedule2Sheet))
            {
                Console.WriteLine("Filling out schedule 2");

                AmountTotal clientTotal = new AmountTotal();

                var currentRow = 8;
                foreach (var invoice in client.Invoices.Values)
                {
                    var dateCell = SetCellValue(schedule2Sheet, _schedule2ColumnStyles, currentRow, "B", invoice.Date);
                    dateCell.SetDataType(XLCellValues.DateTime);

                    var invoiceTotal = invoice.GetTotal();

                    SetCellValue(schedule2Sheet, _schedule2ColumnStyles, currentRow, "C", invoice.InvoiceId);
                    SetCellValue(schedule2Sheet, _schedule2ColumnStyles, currentRow, "E", invoiceTotal.Amount);
                    SetCellValue(schedule2Sheet, _schedule2ColumnStyles, currentRow, "H", invoiceTotal.Assesed);

                    clientTotal.Amount += invoiceTotal.Amount;
                    clientTotal.Assesed += invoiceTotal.Assesed;

                    currentRow++;
                }

                currentRow++;

                SetCellValue(schedule2Sheet, _schedule2ColumnStyles, currentRow, "C", "Grand Total");

                var totalAmountCell = SetCellValue(schedule2Sheet, _schedule2ColumnStyles, currentRow, "E", clientTotal.Amount);
                totalAmountCell.Style.Font.Bold = true;

                var totalAssessedCell = SetCellValue(schedule2Sheet, _schedule2ColumnStyles, currentRow, "H", clientTotal.Assesed);
                totalAssessedCell.Style.Font.Bold = true;
            }
        }

        static void CreateSchedule3(XLWorkbook workbook, AssesedClient client)
        {
            IXLWorksheet schedule3Sheet;
            if (workbook.TryGetWorksheet("Schedule 3", out schedule3Sheet))
            {
                Console.WriteLine("Filling out schedule 3");

                var currentRow = 9;

                var uniqueItems = new HashSet<string>();

                foreach (var invoice in client.Invoices.Values)
                {
                    foreach (var currentEntry in invoice.Entries)
                    {
                        if (!uniqueItems.Contains(currentEntry.ItemId))
                        {
                            uniqueItems.Add(currentEntry.ItemId);

                            SetCellValue(schedule3Sheet, _schedule3ColumnStyles, currentRow, "B", currentEntry.ItemClientId);
                            SetCellValue(schedule3Sheet, _schedule3ColumnStyles, currentRow, "C", currentEntry.ItemId);
                            SetCellValue(schedule3Sheet, _schedule3ColumnStyles, currentRow, "D", currentEntry.ItemDescription);

                            currentRow++;
                        }
                    }
                }
            }
        }

        static IXLCell SetCellValue<T>(IXLWorksheet sheet, Dictionary<string, IXLStyle> styleMap, int row, string column, T value)
        {
            var cell = sheet.Cell(row, column);
            if (!styleMap.ContainsKey(column))
            {
                styleMap[column] = cell.Style;
            }
            cell.Style = styleMap[column];

            cell.SetValue(value);

            return cell;
        }
    }
}