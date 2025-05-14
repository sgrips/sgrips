using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sgrips.Framework.FileSystemDatabase
{
    public class FieldAttributes : Attribute
    {

        public FieldAttributes() {
        }

        public bool Encrypt { get; set; }
    }
}
