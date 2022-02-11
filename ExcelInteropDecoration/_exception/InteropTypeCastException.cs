using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelInteropDecoration._exception
{
    public class InteropTypeCastException : Exception
    {
        public InteropTypeCastException(string message) : base(message)
        {
        }

        public InteropTypeCastException(Type expected, Type actual) :
            this($"Unexpected type encountered. Expected a type of '{expected.FullName}' but found '{actual.FullName}'")
        {
        }
    }
}
