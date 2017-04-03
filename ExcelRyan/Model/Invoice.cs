using System;
using System.Collections.Generic;

namespace ExcelRyan.Model
{
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

}