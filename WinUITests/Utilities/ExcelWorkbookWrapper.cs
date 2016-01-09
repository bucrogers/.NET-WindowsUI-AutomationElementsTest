using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

using Excel = Microsoft.Office.Interop.Excel;

namespace ItemSystemTests.Utilities
{
    public class ExcelWorkbookWrapper : IDisposable 
    {
        private Excel.Workbook _wb;

        public ExcelWorkbookWrapper(Excel.Workbook wb)
        {
            _wb = wb;
        }

        public Excel.Workbook Wb { get { return _wb;  } }

        public void Dispose()
        {
            //TODO: Implement idisposable finalizers pattern if warranted
            _wb.Close(false); //saveChanges
            while (Marshal.ReleaseComObject(_wb) > 0) { }
            _wb = null;
            GC();
        }

        private static void GC()
        {
            System.GC.Collect();
            System.GC.WaitForPendingFinalizers();
            System.GC.Collect();
            System.GC.WaitForPendingFinalizers();
        }
    }

}
