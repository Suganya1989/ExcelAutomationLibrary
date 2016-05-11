using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAutomationLibrary
{
    public class ExcelReader<T>
    {
        public List<T> ReadFile(Byte[] filebytes)
        {
            var objList = new List<T>();
            objList = new ExcelReaderFactory<T>().ReadExcelContents(filebytes);
            return objList;
        }
    }
}
