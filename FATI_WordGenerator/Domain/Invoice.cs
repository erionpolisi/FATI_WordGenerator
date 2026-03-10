using ClosedXML.Excel;
using DocumentFormat.OpenXml.ExtendedProperties;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FATI_WordGenerator.Domain
{
    public class Invoice
    {
        public string? Details { get; set; }
        public string Address { get; set; } = string.Empty;
        public string Period { get; set; } = string.Empty;
        public Company Company { get; set; } = new();
        public List<InvoicePosition> Positions { get; set; } = new();

        public double Sum => Positions.Sum(p => p.Total);
        public double Discount => Sum * 0.03;
        public double Total => Sum - Discount;
    }
}
