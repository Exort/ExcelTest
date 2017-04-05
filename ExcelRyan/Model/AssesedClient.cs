using System;
using System.Collections.Generic;

namespace ExcelRyan.Model
{
    class AssesedClient
    {
        public string Id;
        public string Name;
        public string Address;

        public AmountTotal LastCalculatedTotal;

        public Dictionary<string, Invoice> Invoices = new Dictionary<string, Invoice>();

        public AssesedClient(string id)
        {
            Id = id;
        }

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
}