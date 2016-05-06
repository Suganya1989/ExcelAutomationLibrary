using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAutomationLibrary
{
    [AttributeUsage(AttributeTargets.Property, Inherited = false)]
    public class HeaderAttribute: Attribute
    {
        private string _headerName;
        private int _position;
        public HeaderAttribute(string headerName , int position)
        {
            this._headerName = headerName;
            this._position = position;
        }
        public string HeaderName {
            get {return _headerName;}
            set { _headerName = value;}
        }

        public int Positon {
            get{return _position;}
            set { _position = value; }
        }
    }
}
