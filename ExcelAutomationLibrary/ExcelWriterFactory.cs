using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAutomationLibrary
{
    public class ExcelWriterFactory<TInput>
    {
        public SpreadsheetDocument CreateDocument(IEnumerable<TInput> Input, IEnumerable<Header> headers)
        {
            SharedStringTablePart sharedStringTablePart;
            WorkbookStylesPart workbookStylesPart;
            var openXMLElementList = new List<OpenXmlElement>();

            using (SpreadsheetDocument document = SpreadsheetDocument.Create("ExcelFile.xlsx", SpreadsheetDocumentType.Workbook))
            {
                // Add a WorkbookPart to the document.
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();
                // Add a WorksheetPart to the WorkbookPart.
                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                // Shared string table
                sharedStringTablePart = document.WorkbookPart.AddNewPart<SharedStringTablePart>();
                sharedStringTablePart.SharedStringTable = new SharedStringTable();
                sharedStringTablePart.SharedStringTable.Save();

                // Stylesheet
                workbookStylesPart = document.WorkbookPart.AddNewPart<WorkbookStylesPart>();
                workbookStylesPart.Stylesheet = new Stylesheet();
                workbookStylesPart.Stylesheet.Save();

                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

                Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet 1" };

                sheets.Append(sheet);
                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

             
                int index = 1;
               
                Row row;
                row = new Row() { RowIndex = 1 };

                headers.OrderBy(m=> m.Position).ToList().ForEach(m => openXMLElementList.Add(ConstructCell(m.HeaderName, CellValues.String)));

                row.Append(openXMLElementList.ToArray<OpenXmlElement>());

                sheetData.Append(row);
                
                foreach (var record in Input)
                {
                    var rowIndex = index + 1;
                    row = new Row() { RowIndex = (UInt32)rowIndex };
                   
                    var props = record.GetType().GetProperties().OrderBy(m => ((m.GetCustomAttributes(true)[0]).GetType().GetProperty("Positon"))).ToList();
                    
                   
                    var dataOpenXMLElement = new List<OpenXmlElement>();

                    foreach (var prop in props)
                    {
                       dataOpenXMLElement.Add(ConstructCell(record.GetType().GetProperty(prop.Name).GetValue(record, null).ToString(),CellValues.String));
                    }
                    row.Append(dataOpenXMLElement);
                   
                    


                    index = index + 1;
                    sheetData.Append(row);
                }

                workbookPart.Workbook.Save();
                worksheetPart.Worksheet.Save();

                return document;

            }

            return null;
        }

        private Cell ConstructCell(string value, CellValues dataType)
        {
            return new Cell()
            {
                CellValue = new CellValue(value),
                DataType = new EnumValue<CellValues>(dataType)

            };
        }

    }
}
