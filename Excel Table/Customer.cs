using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Table
{
    class Customer
    {
        public int Id { get; set; }
        public string NameOfTheOrganization { get; set; }
        public string Address { get; set; }
        public string Name { get; set; }
        public Product OrderedProduct { get; set; }
        public List<Order> Orders { get; } = new List<Order>();
    }
}
