using CsharpExtras.Extensions;
using InteropDecoration.Decorator._base;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InteropDecoration.Decorator.names
{
    class NamesDImpl : DecoratorBase, INamesD
    {
        public Names RawNames { get; }

        public NamesDImpl(IInteropDAPI api, Names names) : base(api)
        {
            RawNames = names ?? throw new ArgumentNullException(nameof(names));
        }

        public INameD? this[string index] => NameOrNull(index);
                
        private INameD? NameOrNull(string index)
        {
            if (string.IsNullOrWhiteSpace(index))
            {
                Log.Debug("Empty name passed to Names object. Returning null.");
                return null;
            }
            foreach (Name name in RawNames)
            {
                if(Normalize(name.Name) == Normalize(index))
                {
                    return DecoratorFactory.NameD(name);
                }
            }
            Log.Debug($"The name {index} was not found in this Names object. Returning null.");
            return null;
        }

        private string Normalize(string name)
        {
            return name.RemoveWhitespace().ToLower();
        }

        public void DeleteNamedRange(string nameToDelete)
        {
            try
            {
                INameD? name = NameOrNull(nameToDelete);
                if(name == null)
                {
                    Log.Info($"Name {nameToDelete} not found. No deletion done.");
                    return;
                }
                name.Delete();
            }
            catch (Exception ex)
            {
                Log.Warn(string.Format("Exception while deleting named range '{0}'", nameToDelete), ex);
            }
        }

        public bool DeleteNamedRangeSet(ISet<string> setOfNamesToDelete)
        {
            if (setOfNamesToDelete.Count == 0)
            {
                Log.Debug("Emtpy set of names means nothing to delete, so delete has trivially succeeded.");
                return true;
            }

            bool allDeleted = true;
            try
            {
                ISet<string> normalizedSet = NormalizedNamesCopy(setOfNamesToDelete);
                foreach (Name name in RawNames)
                {
                    if (normalizedSet.Contains(Normalize(name.Name)))
                    {
                        try
                        {
                            name.Delete();
                        }
                        catch (Exception ex)
                        {
                            Log.Warn(string.Format("Exception while deleting named range '{0}'", name.Name), ex);
                            allDeleted = false;
                        }
                    }
                }
                return allDeleted;
            }
            catch (Exception ex)
            {
                Log.Warn(string.Format("Exception while deleting a set of named ranges"), ex);
                return false;
            }
        }

        private ISet<string> NormalizedNamesCopy(ISet<string> names)
        {
            ISet<string> normalizedCopy = new HashSet<string>();
            foreach(string name in names)
            {
                normalizedCopy.Add(Normalize(name));
            }
            return normalizedCopy;
        }

        public IEnumerator<INameD> GetEnumerator()
        {
            foreach(object rawObject in RawNames)
            {
                if(rawObject is Name rawName)
                {
                    INameD name = DecoratorFactory.NameD(rawName);
                    yield return name;
                }
            }
        }

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
