using ExcelInteropDecoration.Decorator._base;
using Microsoft.Office.Interop.Excel;

namespace ExcelInteropDecoration.Decorator.listObjects
{
    class ListObjectsDImpl : DecoratorBase, IListObjectsD
    {
        public ListObjectsDImpl(IInteropDAPI api, ListObjects listObjects) : base(api)
        {
            RawListObjects = listObjects;
        }

        public ListObjects RawListObjects { get; }

        public bool HasTable(string tableName)
        {
            try
            {
                ListObject x = RawListObjects[tableName];
                return true;
            }
            catch (Exception)
            {
                Log.Debug("List objects indexer returned false => has table is false");
                return false;
            }
        }

        public IListObjectD this[string tableName]
        {
            get
            {
                if (!HasTable(tableName))
                {
                    throw new IndexOutOfRangeException(string.Format(
                        "Not table found with name {0}", tableName));
                }
                return new ListObjectDImpl(InteropDApi, RawListObjects[tableName]);
            }
        }


    }
}
