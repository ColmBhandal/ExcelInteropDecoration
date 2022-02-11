using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelInteropDecoration.Helper.TempState.ComAddin
{
    public interface IComAddinState
    {
        string AddinId { get; set; }
        bool IsConnected { get; set; }
    }
}
