using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Table
{
    class Order
    {
        public int Id { get; set; }
        public int ProductID { get; set; }
        public int CustomerID { get; set; }
        public int orderId { get; set; }
        public int Quantity { get; set; }
        public DateTime Date { get; set; }
    }
}
