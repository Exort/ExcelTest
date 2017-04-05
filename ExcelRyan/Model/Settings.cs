using System;
using System.Collections.Generic;

namespace ExcelRyan.Model
{
    [Serializable]
    public class Settings
    {
        public class EntriesWorksheetColumns
        {
            public int FirstRow;
            public string ClientId;
            public string Date;
            public string InvoiceId;
            public string ItemClientId;
            public string ItemId;
            public string ItemDescription;
            public string Amount;
            public string Assessed;
        }

        public string InputDocumentPath;

        public string ClientInfoDocumentPath;
        public int ClientInfoFirstRow;

        public string SchedulesTemplateDocumentPath;
        public string InvoiceTemplateDocumentPath;

        public string OutputFolder;

        public Dictionary<string, EntriesWorksheetColumns> Worksheets = new Dictionary<string, EntriesWorksheetColumns>();
    }
}