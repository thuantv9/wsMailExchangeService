using System;
using System.Globalization;
using System.IO;
using System.Reflection;

namespace wsEmailExchange
{
    internal class Common
    {
        private static string AppDir = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
        private static string LogPath = Path.Combine(Common.AppDir, "ErrLog.txt");
        private static string BreakLine = "==================================================";

        public static void GhiLog(string Title, string ErrContent, string PreMess = "")
        {
            try
            {
                File.AppendAllText(Common.LogPath.Replace("ErrLog.txt", "ErrLog_" + DateTime.Now.ToString("yyyyMMdd")) + ".txt", Environment.NewLine + Common.BreakLine + Environment.NewLine + "Title: " + Title + Environment.NewLine + "Date: " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") + Environment.NewLine + "Message: " + ErrContent + Environment.NewLine + "More Info: " + PreMess);
            }
            catch
            {
            }
        }

        public static void GhiLog(string Title, Exception ex, string PreMess = "")
        {
            try
            {
                File.AppendAllText(Common.LogPath.Replace("ErrLog.txt", "ErrLog_" + DateTime.Now.ToString("yyyyMMdd")) + ".txt", Environment.NewLine + Common.BreakLine + Environment.NewLine + "Title: " + Title + Environment.NewLine + "Date: " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") + Environment.NewLine + "Message: " + ex.Message + Environment.NewLine + "StackTrace: " + ex.StackTrace + (ex.InnerException == null ? "" : Environment.NewLine + "Inner: " + (object)ex.InnerException) + (PreMess == "" ? "" : Environment.NewLine + "More Info: " + PreMess));
            }
            catch
            {
            }
        }

        public static string XoaNhay(string inputSrt)
        {
            return inputSrt.Replace("'", "''");
        }

        public static DateTime? ValidDate(object val)
        {
            try
            {
                if (val is DateTime)
                    return new DateTime?((DateTime)val);
                return new DateTime?(DateTime.ParseExact(val.ToString(), "yyyy-MM-dd HH:mm:ss", (IFormatProvider)CultureInfo.InvariantCulture));
            }
            catch
            {
                return new DateTime?();
            }
        }

        public static Decimal? ValidDecimal(object val)
        {
            try
            {
                return new Decimal?(Decimal.Parse(val.ToString()));
            }
            catch
            {
                return new Decimal?();
            }
        }
    }
}
