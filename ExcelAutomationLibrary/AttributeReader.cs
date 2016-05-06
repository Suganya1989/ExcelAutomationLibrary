using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;

namespace ExcelAutomationLibrary
{
    public class AttributeReader
    {
        /// <summary>
        /// Gets all custom attributes to get headers.
        /// </summary>
        /// <param name="input">The input.</param>
        /// <returns></returns>
        public HeaderDetails GetAllAttributes(object input)
        {
            var headerDetails = new HeaderDetails();
            var attributeList = new List<object>();
            var headers = new List<Header>();
            var list = input.GetType().GetProperties();
            list.ToList().ForEach(m => attributeList.Add(m.GetCustomAttributes(true)));
            foreach(var obj in attributeList)
            {
                var header = new Header();
                header.HeaderName = ((HeaderAttribute)(((object[])(obj))[0])).HeaderName;
                header.Position = ((HeaderAttribute)(((object[])(obj))[0])).Positon;
                headers.Add(header);
             }
            headerDetails.Headers = headers;
           return headerDetails;
        }

        
    }
}
