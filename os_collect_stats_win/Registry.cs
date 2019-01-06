using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace os_collect_stats_win
{
    class Registry
    {
        public Registry() {}


        public static Microsoft.Win32.RegistryKey GetRegistryKey(string sKey)
        {
           return Microsoft.Win32.Registry.LocalMachine.OpenSubKey(sKey);
        }

        public static Object GetRegistryValue(string sKey, string sValue)
        {
            Object obj = null;

            using (Microsoft.Win32.RegistryKey key = GetRegistryKey(sKey))
            {
                if (key != null)
                {
                    obj = key.GetValue(sValue);
                }
            }
            
            return obj;
        }
    }
}
