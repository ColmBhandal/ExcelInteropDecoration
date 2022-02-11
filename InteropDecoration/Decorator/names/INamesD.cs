using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace InteropDecoration.Decorator.names
{
    public interface INamesD : IEnumerable<INameD>
    {
        /// <summary>
        /// Gets a given name by index
        /// </summary>
        /// <param name="index">A string value representing the name</param>
        /// <returns><returns>The Name object with the given name. Returns null if the index passed in is not a known name</returns>
        INameD? this[string index] { get; }

        Names RawNames { get; }

        /// <summary>
        /// Delete the provided named range if it exists.
        /// </summary>
        void DeleteNamedRange(string nameToDelete);

        /// <summary>
        /// Delete all named ranges in the provided set.
        /// </summary>
        /// <param name="setOfNamesToDelete"></param>
        /// <returns>true if all provided named ranges were deleted, false if any failed</returns>
        bool DeleteNamedRangeSet(ISet<string> setOfNamesToDelete);
    }
}