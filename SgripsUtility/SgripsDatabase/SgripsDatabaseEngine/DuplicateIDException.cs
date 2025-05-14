using System;
using System.Runtime.Serialization;

namespace Sgrips.Framework.FileSystemDatabase
{
    [Serializable]
    internal class DuplicateIDException : Exception
    {
        public DuplicateIDException()
        {
        }

        public DuplicateIDException(string message) : base(message)
        {
        }

        public DuplicateIDException(string message, Exception innerException) : base(message, innerException)
        {
        }

        protected DuplicateIDException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }
    }
}