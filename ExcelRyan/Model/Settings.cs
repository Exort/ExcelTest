﻿using System;
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
}