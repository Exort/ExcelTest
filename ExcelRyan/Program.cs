using System;
using System.CodeDom;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using ClosedXML.Excel;
using Newtonsoft.Json;

namespace ExcelRyan
{
    internal class Program
    {
        [Serializable]
        public class Settings
        {
            public class WorksheetColumns
            {
                public int FirstRow;
                public string ClientId;
                public string Date;
                public string InvoiceId;
                public string ItemId;
                public string ItemDescription;
                public string Amount;
                public string Assessed;
            }

            public string InputDocumentPath;
            public string TemplateDocumentPath;
            public string OutputFolder;

            public Dictionary<string, WorksheetColumns> Worksheets = new Dictionary<string, WorksheetColumns>();
        }

        class AmountTotal
        {
            public double Amount;
            public double Assesed;
        }

        class InvoiceEntry
        {
            public string ClientId;
            public string Date;
            public string InvoiceId;
            public string ItemId;
            public string ItemDescription;
            public double Amount;
            public double Assesed;
        }

        class Invoice
        {
            public string ClientId;
            public string Date;
            public string InvoiceId;

            public List<InvoiceEntry> Entries = new List<InvoiceEntry>();

            public void AddEntry(InvoiceEntry entry)
            {
                if (Entries.Contains(entry))
                {
                    throw new Exception("Trying to add same entry twice");
                }

                Entries.Add(entry);
            }

            public AmountTotal GetTotal()
            {
                var total = new AmountTotal();
                foreach (var invoiceEntry in Entries)
                {
                    total.Amount += invoiceEntry.Amount;
                    total.Assesed += invoiceEntry.Assesed;
                }
                return total;
            }
        }

        class AssesedClient
        {
            public string Id;

            public Dictionary<string, Invoice> Invoices = new Dictionary<string, Invoice>();

            public void AddEntry(InvoiceEntry entry)
            {
                if (!Invoices.ContainsKey(entry.InvoiceId))
                {
                    AddInvoice(new Invoice
                    {
                        ClientId = Id,
                        InvoiceId = entry.InvoiceId,
                        Date = entry.Date
                    });
                }

                Invoices[entry.InvoiceId].AddEntry(entry);
            }

            private void AddInvoice(Invoice invoice)
            {
                if (Invoices.ContainsKey(invoice.InvoiceId))
                {
                    throw new Exception("Trying to add same invoice twice");
                }

                Invoices[invoice.InvoiceId] = invoice;
            }
        }

        private const string TestDataPath = @"D:\Projects\ExcelTest\ExcelRyan\XLSFiles\Test.xlsx";

        private const string RawDataPath =
            @"D:\Projects\ExcelTest\ExcelRyan\XLSFiles\KF & KFDI Back-up file for Ryan - what left.xlsx";

        private const string TemplatePath = @"D:\Projects\ExcelTest\ExcelRyan\XLSFiles\Client Template.xlsx";
        private const string PackagePathFormat = @"D:\Projects\ExcelTest\ExcelRyan\XLSFiles\{0}.xlsx";

        private static List<InvoiceEntry> _allEntries = new List<InvoiceEntry>();
        private static Dictionary<string, AssesedClient> _clients = new Dictionary<string, AssesedClient>();
        private static Settings _settings;

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
#if !DEBUG
                Console.WriteLine("Press any key to close");
                Console.ReadKey();
#endif
            }
        }

        static Dictionary<string, IXLStyle> _schedule2ColumnStyles = new Dictionary<string, IXLStyle>();
        static Dictionary<string, IXLStyle> _schedule3ColumnStyles = new Dictionary<string, IXLStyle>();

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
            Console.WriteLine($"Loading workbook {RawDataPath}");
            Console.WriteLine("This might take a while ^^");

            using (var workbook = new XLWorkbook(RawDataPath))
            {
                Console.WriteLine($"{RawDataPath} loaded.");

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

            var filePath = String.Format(PackagePathFormat, clientID);

#if !DEBUG
            if (File.Exists(filePath))
            {
                throw new FileLoadException($"{filePath} already exists");
            }
#endif

            File.Copy(TemplatePath, filePath, true);

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
                    var dateCell = SetCellValue(schedule2Sheet, currentRow, "B", invoice.Date);
                    dateCell.SetDataType(XLCellValues.DateTime);

                    var invoiceTotal = invoice.GetTotal();

                    SetCellValue(schedule2Sheet, currentRow, "C", invoice.InvoiceId);
                    SetCellValue(schedule2Sheet, currentRow, "E", invoiceTotal.Amount);
                    SetCellValue(schedule2Sheet, currentRow, "H", invoiceTotal.Assesed);

                    clientTotal.Amount += invoiceTotal.Amount;
                    clientTotal.Assesed += invoiceTotal.Assesed;

                    currentRow++;
                }

                currentRow++;

                SetCellValue(schedule2Sheet, currentRow, "C", "Grand Total");

                var totalAmountCell = SetCellValue(schedule2Sheet, currentRow, "E", clientTotal.Amount);
                totalAmountCell.Style.Font.Bold = true;

                var totalAssessedCell = SetCellValue(schedule2Sheet, currentRow, "H", clientTotal.Assesed);
                totalAssessedCell.Style.Font.Bold = true;
            }
        }

        static void CreateSchedule3(XLWorkbook workbook, AssesedClient client)
        {
            IXLWorksheet schedule2Sheet;
            if (workbook.TryGetWorksheet("Schedule 3", out schedule2Sheet))
            {
                Console.WriteLine("Filling out schedule 3");

                var currentRow = 9;

                foreach (var invoice in client.Invoices.Values)
                {
                    currentRow++;
                }
                currentRow++;
            }
        }

        static IXLCell SetCellValue<T>(IXLWorksheet sheet, int row, string column, T value)
        {
            var cell = sheet.Cell(row, column);
            if (!_schedule2ColumnStyles.ContainsKey(column))
            {
                _schedule2ColumnStyles[column] = cell.Style;
            }
            cell.Style = _schedule2ColumnStyles[column];

            cell.SetValue(value);

            return cell;
        }
    }
}