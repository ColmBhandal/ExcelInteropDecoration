using CsharpExtras.Api;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelInteropDecoration._base
{
    internal class BaseClass
    {
        protected IInteropDAPI InteropDApi { get; }

        public BaseClass(IInteropDAPI interopDApi)
        {
            InteropDApi = interopDApi ?? throw new ArgumentNullException(nameof(interopDApi));
        }

        private ILogger? _log;
        public ILogger Log => _log ??= InteropDApi.NewLogger(GetType());
        private ICsharpExtrasApi? _csharpExtrasApi;
        public ICsharpExtrasApi CsharpExtrasApi => _csharpExtrasApi ??= new CsharpExtrasApi();
    }
}
