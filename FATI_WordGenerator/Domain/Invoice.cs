using DocumentFormat.OpenXml.ExtendedProperties;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FATI_WordGenerator.BusinessLayer
{
    public class Invoice
    {
        public string Address { get; set; }
        public string Period { get; set; }
        public Company Company { get; set; }
        public List<InvoicePosition> Positions { get; set; } = new();

        public double Sum => Positions.Sum(p => p.Total);
        public double Discount => Sum * 0.03;
        public double Total => Sum - Discount;
    }
}
