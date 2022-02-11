using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelInteropDecoration.Helper.TempState.ComAddin
{
    internal class ComAddinStatePoco : IComAddinState
    {
        public string AddinId { get; set; } = "";
        public bool IsConnected { get; set; }
    }
}
