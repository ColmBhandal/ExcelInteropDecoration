using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InteropDecoration._base
{
    public interface IBuilder<T>
    {
        T Build();
    }
}
