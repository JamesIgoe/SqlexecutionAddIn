using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace SQLExecution
{
    public static class ComCleanup
    {
        /// <summary>
        ///     sets COM object count to zero and sets to null for GC
        /// </summary>
        /// <param name="o"></param>
        public static void FinalReleaseAndNull(object o)
        {
            try
            {
                Marshal.FinalReleaseComObject(o);
            }
            catch
            {
            }
            finally
            {
                o = null;
            }
        }

        /// <summary>
        ///     calls garbage collection twice
        /// </summary>
        /// <param name="o"></param>
        public static void GarbageCleanup()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

    }
}
