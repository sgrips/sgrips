using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sgrips.Framework.FileSystemDatabase
{
    class UniqueIDNotFoundException : Exception
    {


        public UniqueIDNotFoundException() : base("Attenzione non è presente la chiave primaria nella tabella")
        {

        }

        public UniqueIDNotFoundException(string message) : base(message )
        {

        }

        public UniqueIDNotFoundException(string message,Exception innerException) : base(message, innerException)
        {

        }
    }
}
