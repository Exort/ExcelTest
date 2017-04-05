using System;
using System.Collections.Generic;

namespace ExcelRyan.Model
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
            public string ItemClientId;
            public string ItemId;
            public string ItemDescription;
            public string Amount;
            public string Assessed;
        }

        public string InputDocumentPath;
        public string SchedulesTemplateDocumentPath;
        public string InvoiceTemplateDocumentPath;

        public string OutputFolder;

        public Dictionary<string, WorksheetColumns> Worksheets = new Dictionary<string, WorksheetColumns>();
    }
}