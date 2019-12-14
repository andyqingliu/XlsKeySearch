using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using log4net;
using log4net.Config;

namespace XlsKeySearch.ExcelHandler
{
    public class Debug
    {
        private static readonly ILog LogInfo = LogManager.GetLogger("XlsKeySearchLogger");

        private static string GetFormatStr()
        {
            string fileName = Util.GetFileName();
            string methodName = Util.GetMethodName();
            int codeLineNum = Util.GetLineNum();
            string formatStr = string.Format("{0}[{1}()]({2})", fileName, methodName, codeLineNum);
            return formatStr;
        }

        public static void Log(string str, params object[] args)
        {
            if (Util.IsEnableLog)
            {
                LogInfo.InfoFormat(GetFormatStr() + str, args);
            }
        }

        public static void Warn(string str, params object[] args)
        {
            if (Util.IsEnableLog)
            {
                LogInfo.WarnFormat(GetFormatStr() + str, args);
            }
        }

        public static void Error(string str, params object[] args)
        {
            if (Util.IsEnableLog)
            {
                LogInfo.ErrorFormat(GetFormatStr() + str, args);
            }
        }
    }
}
