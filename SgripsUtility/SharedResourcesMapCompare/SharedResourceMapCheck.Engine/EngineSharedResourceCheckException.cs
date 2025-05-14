using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sgrips.Utility.IO.SharedResourceMapCheck.Engine
{
    class EngineSharedResourceCheckException : Exception
    {
        public EngineSharedResourceCheckException() { }

        public EngineSharedResourceCheckException(string msg) :
            base(msg)
        { }

        public EngineSharedResourceCheckException(string msg , Exception innerException ) :
            base(msg, innerException)
        { }

    }
}
