using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
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

        static Dictionary<string, IXLStyle> _schedule2Styles = new Dictionary<string, IXLStyle>();
        static Dictionary<string, IXLStyle> _schedule3Styles = new Dictionary<string, IXLStyle>();

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

            FilloutClientInformation();

            Console.WriteLine("Data loaded. Processing each clients found");

            foreach (var client in _clients.Values)
            {
                var outputFolder = _settings.OutputFolder;

                Directory.CreateDirectory(outputFolder);

                CreateSchedules(outputFolder, client);
                CreateRyanInvoice(outputFolder, client);
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

                    Settings.EntriesWorksheetColumns sheetSettings;
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
                            ItemId = sheet.Cell(currentRow, sheetSettings.ItemId).GetString(),
                            ItemDescription = sheet.Cell(currentRow, sheetSettings.ItemDescription).GetString(),
                            Amount = sheet.Cell(currentRow, sheetSettings.Amount).GetDouble(),
                            Assesed = sheet.Cell(currentRow, sheetSettings.Assessed).GetDouble(),
                        };

                        _allEntries.Add(invoiceEntry);

                        if (!_clients.ContainsKey(invoiceEntry.ClientId))
                        {
                            _clients[invoiceEntry.ClientId] = new AssesedClient(invoiceEntry.ClientId);
                        }
                        _clients[invoiceEntry.ClientId].AddEntry(invoiceEntry);

                        currentRow++;
                    }

                    Console.WriteLine($"{sheet.Name} done. Processed {currentRow - sheetSettings.FirstRow + 1} rows.");
                }
            }
            return _allEntries;
        }

        static void FilloutClientInformation()
        {
            Console.WriteLine($"Loading workbook {_settings.ClientInfoDocumentPath}");
            Console.WriteLine("This might take a while ^^");

            using (var workbook = new XLWorkbook(_settings.ClientInfoDocumentPath))
            {
                Console.WriteLine($"{_settings.ClientInfoDocumentPath} loaded.");

                foreach (var sheet in workbook.Worksheets)
                {
                    Console.WriteLine($"Processing sheet {sheet.Name}");

                    var rowCount = sheet.RowCount();
                    var currentRow = _settings.ClientInfoFirstRow;

                    while (currentRow <= rowCount && !string.IsNullOrEmpty(sheet.Cell(currentRow, "A").GetString()))
                    {
                        Console.WriteLine($"Processing row {currentRow}");

                        var clientId = sheet.Cell(currentRow, "A").GetString();

                        AssesedClient assesedClient;
                        if (!_clients.TryGetValue(clientId, out assesedClient))
                        {
                            Console.WriteLine($"Skipping {clientId}, no known invoices for that client");
                            currentRow++;

                            continue;
                        }

                        assesedClient.Name = sheet.Cell(currentRow, "B").GetString();
                        assesedClient.Address = sheet.Cell(currentRow, "C").GetString();
                        assesedClient.City = sheet.Cell(currentRow, "D").GetString();
                        assesedClient.PostalCode = sheet.Cell(currentRow, "E").GetString();
                        assesedClient.RyanInvoiceId = sheet.Cell(currentRow, "F").GetString();

                        currentRow++;

                    }

                    Console.WriteLine($"{sheet.Name} done. Processed {currentRow - _settings.ClientInfoFirstRow + 1} rows.");
                }
            }
        }

        static void CreateSchedules(string outputFolder, AssesedClient client)
        {
            Console.WriteLine($"Creating package for {client.Id}");

            var filePath = String.Format(Path.Combine(outputFolder, $"Schedule {client.Id}.xlsx"));

#if !DEBUG
            if (File.Exists(filePath))
            {
                throw new FileLoadException($"{filePath} already exists");
            }
#endif

            File.Copy(_settings.SchedulesTemplateDocumentPath, filePath, true);

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
                    var dateCell = SetCellValue(schedule2Sheet, _schedule2Styles, currentRow, "B", invoice.Date);
                    dateCell.SetDataType(XLCellValues.DateTime);

                    var invoiceTotal = invoice.GetTotal();

                    SetCellValue(schedule2Sheet, _schedule2Styles, currentRow, "C", invoice.InvoiceId);
                    SetCellValue(schedule2Sheet, _schedule2Styles, currentRow, "E", invoiceTotal.Amount);
                    SetCellValue(schedule2Sheet, _schedule2Styles, currentRow, "H", invoiceTotal.Assesed);

                    clientTotal.Amount += invoiceTotal.Amount;
                    clientTotal.Assesed += invoiceTotal.Assesed;

                    currentRow++;
                }

                client.LastCalculatedTotal = clientTotal;

                currentRow++;

                SetCellValue(schedule2Sheet, _schedule2Styles, currentRow, "C", "Grand Total");

                var totalAmountCell = SetCellValue(schedule2Sheet, _schedule2Styles, currentRow, "E",
                    clientTotal.Amount);
                totalAmountCell.Style.Font.Bold = true;

                var totalAssessedCell = SetCellValue(schedule2Sheet, _schedule2Styles, currentRow, "H",
                    clientTotal.Assesed);
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

                List<InvoiceEntry> allEntries = new List<InvoiceEntry>();

                foreach (var invoice in client.Invoices.Values)
                {
                    foreach (var currentEntry in invoice.Entries)
                    {
                        if (string.IsNullOrEmpty(currentEntry.ItemId) || uniqueItems.Contains(currentEntry.ItemId))
                        {
                            continue;
                        }
                        uniqueItems.Add(currentEntry.ItemId);
                        allEntries.Add(currentEntry);
                    }
                }

                allEntries = allEntries.OrderBy(x => x.ItemDescription).ToList();

                foreach (var currentEntry in allEntries)
                {
                    SetCellValue(schedule3Sheet, _schedule3Styles, currentRow, "B", currentEntry.ClientId);
                    SetCellValue(schedule3Sheet, _schedule3Styles, currentRow, "C", currentEntry.ItemId);
                    SetCellValue(schedule3Sheet, _schedule3Styles, currentRow, "D", currentEntry.ItemDescription);

                    currentRow++;
                }
            }
        }

        static void CreateRyanInvoice(string outputFolder, AssesedClient client)
        {
            Console.WriteLine($"Creating invoice for {client.Id}");

            var filePath = Path.Combine(outputFolder, $"Invoice {client.Id}.xlsx");

#if !DEBUG
            if (File.Exists(filePath))
            {
                throw new FileLoadException($"{filePath} already exists");
            }
#endif

            File.Copy(_settings.InvoiceTemplateDocumentPath, filePath, true);

            using (var workbook = new XLWorkbook(filePath))
            {
                IXLWorksheet invoiceSheet;
                if (workbook.TryGetWorksheet("Sheet1", out invoiceSheet))
                {
                    invoiceSheet.Cell(5, "F").SetValue(client.Id);
                    invoiceSheet.Cell(10, "A").SetValue(client.Name);
                    invoiceSheet.Cell(11, "A").SetValue(client.Address);
                    invoiceSheet.Cell(12, "A").SetValue(client.City);
                    invoiceSheet.Cell(13, "A").SetValue(client.PostalCode);
                    invoiceSheet.Cell(4, "F").SetValue(client.RyanInvoiceId);

                    var dateCell = invoiceSheet.Cell(3, "F");
                    dateCell.SetValue(DateTime.Today.ToString("M/d/yyyy"));
                    dateCell.SetDataType(XLCellValues.DateTime);

                    var dueDateCell = invoiceSheet.Cell(6, "F");
                    dueDateCell.SetValue(DateTime.Today.AddDays(30).ToString("M/d/yyyy"));
                    dueDateCell.SetDataType(XLCellValues.DateTime);

                    invoiceSheet.Cell(16, "F").SetValue(client.LastCalculatedTotal.Assesed);
                    invoiceSheet.Cell(31, "F").SetValue(client.LastCalculatedTotal.Assesed);
                }

                workbook.Save(true);
            }
        }

        static IXLCell SetCellValue<T>(IXLWorksheet sheet, Dictionary<string, IXLStyle> styleMap, int row,
            string column, T value)
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