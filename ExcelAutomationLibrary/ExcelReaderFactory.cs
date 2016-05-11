using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAutomationLibrary
{
    public class ExcelReaderFactory<T>
    {
        public List<T> ReadExcelContents(Byte[] input)
        {
            var objList = new List<T>();
            try
            {
                int emptyRowCount = 0;
                int rowsComplete = 0;

                Stream stream = new MemoryStream(input);

                using (SpreadsheetDocument myDoc = SpreadsheetDocument.Open(stream, true))
                {
                    WorkbookPart workBookPart = myDoc.WorkbookPart;

                    Sheet s = workBookPart.Workbook.Descendants<Sheet>().First();


                    WorksheetPart wsPart = workBookPart.GetPartById(s.Id) as WorksheetPart;


                    Row[] rows = wsPart.Worksheet.Descendants<Row>().ToArray();

                    //assumes the first row contains column names 
                    foreach (Row row in wsPart.Worksheet.Descendants<Row>())
                    {
                        int i = 0;
                        rowsComplete++;

                        bool emptyRow = true;
                        List<object> rowData = new List<object>();
                        string value;

                        foreach (Cell c in row.Elements<Cell>())
                        {
                            value = GetCellValue(c);
                            emptyRow = emptyRow && string.IsNullOrWhiteSpace(value);
                            rowData.Add(value);


                        }


                        if (rowData[1].ToString() == string.Empty && rowData[2].ToString() == string.Empty) // to handle end of records
                        {
                            emptyRowCount++;
                            if (emptyRowCount >= 5)
                            {
                                break;
                            }
                            else
                            {
                                continue;
                            }

                        }
                        if (rowData != null && rowData[1].ToString() != string.Empty && rowsComplete > 1)
                        {
                            Type classInstance = typeof(T);
                            var className = (T)Activator.CreateInstance(classInstance);
                            var properties = className.GetType().GetProperties();
                            foreach (var prop in properties)
                            {
                                className.GetType().GetProperty(prop.Name.ToString()).SetValue(className, rowData[i].ToString());
                                i++;
                            }
                            objList.Add(className);
                        }
                    }


                }

                return objList;
            }

            catch (NotSupportedException ex)
            {
                return objList;
            }
           

        }

        public static string GetCellValue(Cell cell)
        {
            if (cell == null)
                return null;
            if (cell.DataType == null)
                return cell.InnerText;

            string value = cell.InnerText;
            switch (cell.DataType.Value)
            {
                case CellValues.SharedString:
                    // For shared strings, look up the value in the shared strings table.
                    // Get worksheet from cell
                    OpenXmlElement parent = cell.Parent;
                    while (parent.Parent != null && parent.Parent != parent
                            && string.Compare(parent.LocalName, "worksheet", true) != 0)
                    {
                        parent = parent.Parent;
                    }
                    if (string.Compare(parent.LocalName, "worksheet", true) != 0)
                    {
                        throw new Exception("Unable to find parent worksheet.");
                    }

                    Worksheet ws = parent as Worksheet;
                    SpreadsheetDocument ssDoc = ws.WorksheetPart.OpenXmlPackage as SpreadsheetDocument;
                    SharedStringTablePart sstPart = ssDoc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

                    // lookup value in shared string table
                    if (sstPart != null && sstPart.SharedStringTable != null)
                    {
                        value = sstPart.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                    }
                    break;


            }
            return value;
        }
    }
}
