using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sgrips.Framework.FileSystemDatabase
{
    interface ICursor
    {

        bool Read();



        object GetValue(Int32 index);

        string GetName(Int32 index);


        int FieldCount { get;  }

        //Indexer
        object this[int index]
        {
            get;
            set;
        }

        object this[string index]
        {
            get;
            set;
        }

    }
}
