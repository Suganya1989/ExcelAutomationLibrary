using ExcelAutomation.Abstraction;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;

namespace ExcelAutomationLibrary
{
    public class GetHeaders<Tinput> 
           
    {
        public HeaderDetails Execute(Tinput input)
        {
            var headers = new AttributeReader().GetAllAttributes(input);
            return headers;
        }
    }
}
