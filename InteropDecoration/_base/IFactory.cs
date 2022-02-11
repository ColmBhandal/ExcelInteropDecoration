using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InteropDecoration._base
{
    public interface IFactory<T>
    {
        T New();
    }
}
