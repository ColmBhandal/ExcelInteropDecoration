using InteropDecoration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InteropDecorationTest._testBase
{
    public abstract class InteropDTestBase
    {
        private IInteropDAPI? _interopDApi;
        protected IInteropDAPI InteropDApi => _interopDApi ??= new InteropDAPI();
    }
}
