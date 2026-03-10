using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FATI_WordGenerator.Domain
{
    public class Company
    {
        public string Name { get; set; } = string.Empty;
        public string Street { get; set; } = string.Empty;
        public string Number { get; set; } = string.Empty;
        public string PLZ { get; set; } = string.Empty;
        public string City { get; set; } = string.Empty;
        public string ATU { get; set; } = string.Empty;

        public string StreetAndNumber { get; set; } = string.Empty;
        public string PLZAndCity { get; set; } = string.Empty;
    }
}
