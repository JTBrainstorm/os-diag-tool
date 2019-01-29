using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Runtime;

namespace os_collect_stats_win
{
    public class FileLogger
    {
        private static string _tempFolderPath = Path.Combine(Directory.GetCurrentDirectory(), "collect_data");
        private static string _errorDumpFile = Path.Combine(_tempFolderPath, "ErrorDump.txt");


        public static void LogError(string customMessage, string errorMessage)
        {
            File.AppendAllText(_errorDumpFile, DateTime.Now + "\t" + customMessage + "\t" + errorMessage + Environment.NewLine);
        }

    }
}
