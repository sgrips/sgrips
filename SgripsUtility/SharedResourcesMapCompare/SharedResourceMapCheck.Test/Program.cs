using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Sgrips.Utility.IO.SharedResourceMapCheck.Engine;
      

namespace Sgrips.Utility.IO.SharedResourceMapCheck.Test
{
    class Program
    {
        static void Main(string[] args)
        {
            EngineSharedResourceCheck esfc = new EngineSharedResourceCheck();

            foreach (var item in esfc.SharedResourceCheck )
            {
                esfc.RunTask(item);
                esfc.UpdateDateTimeStart(item);  
            }

        }
    }
}
