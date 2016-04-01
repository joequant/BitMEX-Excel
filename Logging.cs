using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;

namespace BitMex
{
    //quick and dirty logging class
    internal class Logging
    {

        static Logging()
        {
        }

        internal static void Log(string format, params object[] args)
        {
            System.Diagnostics.Debug.WriteLine(string.Format(format, args));
        }

    }
}
