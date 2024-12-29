using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using OfficeFileHandler;
using System.Diagnostics;

namespace OfficeFileHandler
{
    public class ExcelFileHandler : IDisposable
    {
        private Excel.Application mExcelApp;
        private Excel.Workbook mWorkbook;

        public ExcelFileHandler(string filePath)
        {
            mExcelApp = new Excel.Application();
            mWorkbook = mExcelApp.Workbooks.Open(filePath);
        }

        public void SetCellValue(int sheetIndex, int row, int column, string value)
        {
            Excel.Worksheet mWorksheet = (Excel.Worksheet)mWorkbook.Sheets[sheetIndex];
            Excel.Range mRange = (Excel.Range)mWorksheet.Cells[row, column];
            mRange.Value = value;

            Marshal.ReleaseComObject(mRange);
            Marshal.ReleaseComObject(mWorksheet);
        }

        public void SetVisible(bool check)
        {
            mExcelApp.Visible = check;
        }

        public Excel.Sheets GetSheets()
        {
            return mWorkbook.Sheets;
        }

        #region Save, Close and Dispose
        public void Save()
        {
            mWorkbook.Save();
        }

        public void Close()
        {
            try
            {
                if (mWorkbook != null)
                {
                    mWorkbook.Close(true); // 저장하고 닫기
                    Marshal.ReleaseComObject(mWorkbook);
                }
            } 
            catch(Exception e)
            {
                Debug.WriteLine(e);
            }

            try
            {
                if (mExcelApp != null)
                {
                    mExcelApp.Quit();
                    Marshal.ReleaseComObject(mExcelApp);
                }
            }
            catch (Exception e)
            {
                Debug.WriteLine(e);
            }
        }

        public void Dispose()
        {
            Close();
        }
        #endregion
    }
}
