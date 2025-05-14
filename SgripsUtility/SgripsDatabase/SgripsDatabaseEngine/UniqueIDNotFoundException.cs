using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sgrips.Framework.FileSystemDatabase
{
    class EntityNotFoundException : Exception
    {


        public EntityNotFoundException() : base("Entità non presente in tabella")
        {

        }

        public EntityNotFoundException(string message) : base(message )
        {

        }

        public EntityNotFoundException(string message,Exception innerException) : base(message, innerException)
        {

        }
    }
}
