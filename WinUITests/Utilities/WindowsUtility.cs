using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ItemSystemTests.Utilities
{
    public static class WindowsUtility
    {
        public static bool WindowsVerSupportsModalInvokeWithoutHang()
        {
            return Win8Server2012OrAbove();
        }

        public static bool Win8Server2012OrAbove()
        {
            return Convert.ToDecimal(Environment.OSVersion.Version.Major + (Environment.OSVersion.Version.Minor / 10M)) >= 6.2M;
        }
    }
}
