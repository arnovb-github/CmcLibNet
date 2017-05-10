using System.Collections.Generic;
using System.Linq;

namespace Vovin.CmcLibNet.Export
{
    /// <summary>
    /// Take a list of strings and returns a list of lists of strings,
    /// in which the total number of characters doesn't exceed a specified length.
    /// </summary>
    internal class ListChopper
    {
        // takes a list and chops it up into separate lists,
        // of which the elements combined are up to x characters long
        private List<List<string>> _choppedList = null;
        private List<string> _listToChop = null;
        private readonly int _maxchars = 0;

        internal ListChopper(List<string> list, int numchars)
        {
            _listToChop = list;
            _maxchars = numchars;
        }

        private void Chop(List<string> list, int maxlength)
        {
            List<string> retval = new List<string>();
            List<string> stillToProcess = list.ToList();
            int count = 0;
            for (int i = 0; i < list.Count; i++)
            {
                count = count + list[i].Length;
                if (count < maxlength)
                {
                    retval.Add(list[i]);
                    stillToProcess.RemoveAt(0);
                }
                else
                {
                    this._choppedList.Add(retval);
                    Chop(stillToProcess, maxlength); // recursive!
                    return;
                }
            }
            this._choppedList.Add(retval); // add remainder
        }

        internal List<List<string>> Portions
        {
            get 
            {
                if (_choppedList == null)
                {
                    _choppedList = new List<List<string>>();
                    this.Chop(_listToChop, _maxchars);
                }
                return _choppedList; 
            }
        }
    }
}
