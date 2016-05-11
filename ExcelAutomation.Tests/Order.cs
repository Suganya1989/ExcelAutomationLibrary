using ExcelAutomationLibrary;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAutomation.Tests
{
    public class Order
    {

        [Header("Item Name", 0)]
        public string Name { get; set; }

        [Header("Item Price", 1)]
        public string Price { get; set; }

        [Header("Discount", 2)]
        public string Experience { get; set; }
    }
}
