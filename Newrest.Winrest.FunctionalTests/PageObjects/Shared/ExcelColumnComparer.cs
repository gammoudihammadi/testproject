using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Shared
{
    public class ExcelColumnComparer : Comparer<string>
    {
        public override int Compare(string x, string y)
        {
            if (x.Length < y.Length)
            {
                return -1;
            }
            if (x.Length > y.Length)
            {
                return 1;
            }
            return string.Compare(x, y, true);
        }
    }
}
