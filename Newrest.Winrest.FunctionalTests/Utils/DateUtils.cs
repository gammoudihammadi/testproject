using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.Utils
{
    /**
     * Un DateUtils.Now qui renvoit une date batch UTC -2 heures
     * 
     * Deprecated - batch du soir
     */
    public static class DateUtils
    {
        public static DateTime Now
        {
            get { return DateTime.Now; }
            //get { return DateTime.Now.AddHours(2); }
        }

        public static bool IsBeforeMidnight()
        {
            return false;
            //return DateTime.Now.Hour >= 24-2;

        }
    }
}
