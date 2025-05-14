using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sgrips.Utility.IO.SharedResourceMapCheck.Engine
{
    public class StringBuilderExtension
    {

        public StringBuilderExtension()
        {
            Buffer = new StringBuilder();
        }

        public Exception Error { get; set; }
        public StringBuilder Buffer { get; set; }


        public bool AppendLine(string value)
        {
            if (Buffer != null) { Buffer.AppendLine(value); return true; }
            else return false;


        }

        public override string ToString()
        {
            return Buffer.ToString();
        }
    }
}
