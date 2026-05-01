using System;
using System.Collections.Generic;

namespace InvoiceEntryApp
{
    internal sealed class PartyInfo
    {
        public string Name { get; set; }
        public string Address { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string PinCode { get; set; }
        public string ContactNo { get; set; }
        public string Tin { get; set; }
    }

    internal sealed class InvoiceItem
    {
        public string ItemNo { get; set; }
        public string Description { get; set; }
        public decimal Quantity { get; set; }
        public decimal Price { get; set; }
        public decimal LineTotal => Quantity * Price;
    }

    internal sealed class InvoiceSubmission
    {
        public PartyInfo Company { get; set; }
        public PartyInfo Customer { get; set; }
        public string InvoiceNo { get; set; }
        public DateTime InvoiceDate { get; set; }
        public string SubTotal { get; set; }
        public string Gst { get; set; }
        public string Total { get; set; }
        public List<InvoiceItem> Items { get; } = new List<InvoiceItem>();
    }
}
