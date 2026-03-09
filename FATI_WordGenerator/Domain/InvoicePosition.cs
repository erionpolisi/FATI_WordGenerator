using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FATI_WordGenerator.Domain
{
    public class InvoicePosition
    {
        public int Quantity { get; set; }
        public string Unit { get; set; }
        public string Description { get; set; }
        public double Price { get; set; }

        public double Total => Quantity * Price;
    }
}
