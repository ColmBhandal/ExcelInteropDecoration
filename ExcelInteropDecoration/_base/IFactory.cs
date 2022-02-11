using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelInteropDecoration._base
{
    public interface IFactory<T>
    {
        T New();
    }
}
