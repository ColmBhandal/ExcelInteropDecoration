using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InteropDecoration._exception
{
    public class InteropObjectGenerationException : Exception
    {
        public InteropObjectGenerationException (string message) : base(message)
        {
        }
        public InteropObjectGenerationException(string message, Exception inner) : base(message, inner)
        {
        }
    }
}
