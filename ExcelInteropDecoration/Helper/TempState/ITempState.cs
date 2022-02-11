
namespace ExcelInteropDecoration.Helper.TempState
{
    public interface ITempState<U>
    {
        U TempObject { get; }

        void RunWithTempOptions(Action action);
    }
}