using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Threading;
using System.Windows.Automation;
using System.Runtime.InteropServices;

using Excel = Microsoft.Office.Interop.Excel;

namespace ItemSystemTests.Utilities
{
    public class ExcelAutoUtility
    {
        private static readonly TimeSpan ExcelOperationsRetryTimeout = TimeSpan.FromSeconds(5);

        public static ExcelAppWrapper OpenAndInstallAddin(string addinName, out dynamic facade)
        {
            var app = new ExcelAppWrapper(new Excel.Application());
            app.App.Visible = true;

            try
            {
                object addInNameObjType = addinName;
                var addin = app.App.COMAddIns.Item(ref addInNameObjType);
                // facade is an object of type IAddInFaçade, but we need to use
                // late binding for everything to work properly.
                facade = addin.Object;

                return app;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                app.Dispose();
                throw;
            }
        }

        public static Excel.Worksheet GetWorkSheetWithTimeout(Excel.Workbook wb, string worksheetName, TimeSpan timeout)
        {
            var sw = Stopwatch.StartNew();
            do
            {
                foreach (Excel.Worksheet wks in wb.Worksheets)
                {
                    if (wks.Name == worksheetName)
                    {
                        return wks;
                 
                    }
                }

                Thread.Sleep(TimeSpan.FromMilliseconds(100)); //be cpu-friendly
            }
            while (sw.Elapsed < timeout);

            throw new TimeoutException("Timed out after " + timeout.TotalSeconds + "' looking for worksheet '" + worksheetName + "'");
        }

        public static AutomationElement GetAppElementFromApp(Excel.Application app)
        {
            //Alternative approach  below - launch exe and get pid
            //http://blog.functionalfun.net/2009/06/introduction-to-ui-automation-with.html

            var el = AutomationElement.FromHandle(new IntPtr(app.Hwnd));

            if (el == null)
            {
                throw new ApplicationException("Could not get Excel AutomationElement from interop instance handle");
            }

            return el;
        }

        public static void SetCellValue(Excel.Worksheet ws, int row, int col, object value)
        {
            Exception thrownEx;
            var sw = Stopwatch.StartNew();
            do
            {
                try
                {
                    ws.Cells[row, col] = value;

                    return;
                }
                catch (Exception ex)
                {
                    thrownEx = ex;

                    Thread.Sleep(TimeSpan.FromMilliseconds(100)); //be cpu-friendly
                }
            }
            while (sw.Elapsed < ExcelOperationsRetryTimeout);

            throw new ApplicationException("Worksheet cell value could not be written after trying for " + ExcelOperationsRetryTimeout.TotalSeconds + "sec",
                    thrownEx);
        }

        public static object GetCellValue(Excel.Worksheet ws, int row, int col)
        {
            Exception thrownEx;
            var sw = Stopwatch.StartNew();
            do
            {
                try
                {
                    return ws.Cells[row, col].Value;
                }
                catch (Exception ex)
                {
                    thrownEx = ex;

                    Thread.Sleep(TimeSpan.FromMilliseconds(100)); //be cpu-friendly
                }
            }
            while (sw.Elapsed < ExcelOperationsRetryTimeout);

            throw new ApplicationException("Worksheet cell value could not be read after trying for " + ExcelOperationsRetryTimeout.TotalSeconds + "sec",
                    thrownEx);
        }

        public static void WaitForNewWorksheetToBeAccessible(Excel.Worksheet ws)
        {
            GetCellValue(ws, 1, 1);
        }

    }
}
