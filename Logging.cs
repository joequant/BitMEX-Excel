using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BitMex
{
    //quick and dirty logging class
    internal class Logging
    {
        private static string _logPath;

        static Logging()
        {
            _logPath = System.IO.Path.GetTempPath() + @"\BitMexRTD.log";
        }

        internal static void Log(string format, params object[] args)
        {
            File.AppendAllText(_logPath, DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff") + " - " + string.Format(format, args) + "\r\n");
        }

    }
}
