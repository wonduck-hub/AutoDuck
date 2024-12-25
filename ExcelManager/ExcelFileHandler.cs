using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace OfficeFileHandler
{
    public class ExcelFileHandler : IDisposable
    {
        private Excel.Application mExcelApp;
        private Excel.Workbook mWorkbook;
        private Excel.Worksheet mWorksheet;
        private Excel.Range mRange;

        public ExcelFileHandler(string filePath)
        {
            mExcelApp = new Excel.Application();
            mWorkbook = mExcelApp.Workbooks.Open(filePath);
        }

        public void SetCellValue(int sheetIndex, int row, int column, string value)
        {
            mWorksheet = (Excel.Worksheet)mWorkbook.Sheets[sheetIndex];
            mRange = (Excel.Range)mWorksheet.Cells[row, column];
            mRange.Value = value;
        }

        public void SetVisible(bool check)
        {
            mExcelApp.Visible = check;
        }

        public void Save()
        {
            mWorkbook.Save();
        }

        public void Dispose()
        {
            if (mRange != null)
            {
                Marshal.ReleaseComObject(mRange);
            }
            if (mWorksheet != null)
            {
                Marshal.ReleaseComObject(mWorksheet);
            }
            if (mWorkbook != null)
            {
                Marshal.ReleaseComObject(mWorkbook);
            }
            if (mExcelApp != null)
            {
                Marshal.ReleaseComObject(mExcelApp);
            }
        }
    }
}