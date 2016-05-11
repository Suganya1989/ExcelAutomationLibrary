using ExcelAutomationLibrary;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Should;

namespace ExcelAutomation.Tests
{
    public class ExcelReaderTests
    {
        public void ShouldBeAbleToReadExcelFile(){

            string fileName = "ExcelFileSample 1.xlsx";
            string filePath = AppDomain.CurrentDomain.BaseDirectory + "\\TestFiles\\" + fileName;

            byte[] bytes = System.IO.File.ReadAllBytes(filePath);
            var excelContents = new ExcelReader<Employee>().ReadFile(bytes);

            excelContents[0].Name.ShouldStartWith("Sugan");
        
        }
    }
}
