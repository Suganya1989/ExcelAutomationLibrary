using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelAutomation.Abstraction;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAutomationLibrary
{
    public class ExcelWriter<TInput>
    {
        public void Write(IEnumerable<TInput> input)
        {
            GetHeaders<TInput> headers = new GetHeaders<TInput>();
            var headerRows = headers.Execute(input.First());

            if (input != null)
            {
                ExcelWriterFactory<TInput> excelFactory = new ExcelWriterFactory<TInput>();
                excelFactory.CreateDocument(input, headerRows.Headers);
            }
         
        }
    }
}
