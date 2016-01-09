using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Threading;

using Excel = Microsoft.Office.Interop.Excel;

namespace ItemSystemTests.Utilities
{
    /// <summary>
    /// Wrapper for excel interop application - ensures process kill on dispose
    /// </summary>
    public class ExcelAppWrapper : IDisposable
    {
        private Excel.Application _app;
        private Process _proc;

        public ExcelAppWrapper(Excel.Application app)
        {
            _app = app;
            _proc = GetExcelProcess(app);
        }

        public Excel.Application App { get { return _app; } }

        public bool IsVersion2013OrAbove()
        {
            return (Convert.ToDecimal(_app.Version) >= 15.0M);
        }

        public bool IsVersion2010OrAbove()
        {
            return (Convert.ToDecimal(_app.Version) >= 14.0M);
        }

        public void Dispose()
        {
            //TODO: Implement idisposable finalizers pattern if warranted
            try
            {
                //First try excel interop to close process
                _app.Quit();

                while (Marshal.ReleaseComObject(_app) > 0) { }
                _app = null;
                GC();

                //If no error thrown by above wait for a while, then process kill if necessary
                //...sometimes the visible excel window is gone but the process is still there

                var waitForQuitToWorkTimeout = TimeSpan.FromSeconds(2);
                var sw = Stopwatch.StartNew();
                do
                {
                    if (_proc.HasExited)
                    {
                        return;
                    }

                    Thread.Sleep(TimeSpan.FromMilliseconds(100)); //be cpu-friendly
                }
                while (sw.Elapsed < waitForQuitToWorkTimeout);

                //Waited up to timeout, proceed with kill
                _proc.Kill();
            }
            catch (Exception)
            {
                //Use process kill if .Quit throws
                if ( ! _proc.HasExited)
                {
                    _proc.Kill();
                }
            }
        }

        private static void GC()
        {
            System.GC.Collect();
            System.GC.WaitForPendingFinalizers();
            System.GC.Collect();
            System.GC.WaitForPendingFinalizers();
        }

        [DllImport("user32.dll")]
        private static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);

        private static Process GetExcelProcess(Excel.Application excelApp)
        {
            int id;
            GetWindowThreadProcessId(excelApp.Hwnd, out id);
            return Process.GetProcessById(id);
        }
}
}
