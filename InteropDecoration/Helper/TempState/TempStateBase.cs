using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InteropDecoration.Helper.TempState
{
    internal abstract class TempStateBase<U> : ITempState<U>
    {
        protected U PermanentObjectMemory { get; }
        public U TempObject { get; }

        U PermanentObject { get; }

        protected TempStateBase(U permanentObject)
        {
            PermanentObjectMemory = DeepCopy(permanentObject);
            TempObject = DeepCopy(permanentObject);
            PermanentObject = permanentObject;
        }

        public void RunWithTempOptions(Action action)
        {
            if (action == null)
            {
                return;
            }
            Copy(TempObject, PermanentObject);
            action.Invoke();
            Copy(PermanentObjectMemory, PermanentObject);
        }

        protected U DeepCopy(U source)
        {
            U u = NewPoco();
            Copy(source, u);
            return u;
        }
        protected abstract U NewPoco();

        /// <summary> Copies various data from one object to the other. We only have to copy the data we care about.</summary>
        /// <param name="from">The source of the copy. This should remain unchanged.</param>
        /// <param name="to">The target of the copy. This will get updated with values from source.</param>
        protected abstract void Copy(U from, U to);
    }
}
