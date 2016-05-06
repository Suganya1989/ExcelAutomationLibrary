using ExcelAutomationLibrary;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAutomation.Tests
{
    public class Employee
    {
        
        [Header("Employee Name", 0)]
        public string Name { get; set; }

        [Header("Emp Id", 1)]
        public string Id { get; set; }

        [Header("Experience", 2)]
        public string Experience { get; set; }

    }
}
