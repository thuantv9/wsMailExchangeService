using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace wsEmailExchange.Util
{
    public class SafeConvert
    {
        public static int ToInt(object obj)
        {
            if (obj == null) return 0;
            try
            {
                return Convert.ToInt32(obj);
            }
            catch
            {
                return 0;
            }
        }
    }
}
