using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using OfficeFileHandler;
using System.Diagnostics;
using System.Data.Common;

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

        public Excel.Sheets GetSheets()
        {
            return mWorkbook.Sheets;
        }

        public void MakeNewSheet(string sheetName)
        {

        }

        public bool MsCetsaRun(int sheetIndex)
        {
            Debug.Assert(sheetIndex > 0);

            Excel.Worksheet worksheet = (Excel.Worksheet)mWorkbook.Sheets[sheetIndex];
            // 가장 작은 인덱스에 있는 값이 있는 셀에서 가장 큰 인덱스에 있는 값이 있는 셀까지가 밤위(빈칸 포함)
            Excel.Range range = worksheet.UsedRange;

            bool isTableFound = false;
            string[] tableHead = { "126", "127", "128", "129", "130", "131" };
            int tableStartRow;
            int tableStartColumn;
            for (int i = 1; i <= range.Rows.Count; ++i)
            {
                int checkCount = 0;
                for (int j = 1; j <= range.Columns.Count; ++j)
                {
                    Excel.Range cell = (Excel.Range)range.Cells[i, j]; 
                    if (cell.Text == tableHead[checkCount])
                    {
                        if (checkCount == 0)
                        {
                            tableStartRow = i + 1;
                            tableStartColumn = j - 1;
                        }
                        if (checkCount == 5)
                        {
                            isTableFound = true;
                            goto IS_FOUND_TABLE;
                        }
                        ++checkCount;
                    }
                    else
                    {
                        checkCount = 0;
                    }
                    Marshal.ReleaseComObject(cell);
                }
            }

            if (!isTableFound)
            {
                goto IS_NOT_FOUND_TABLE;
            }

            IS_FOUND_TABLE:



            IS_NOT_FOUND_TABLE:
            Marshal.ReleaseComObject(range);
            Marshal.ReleaseComObject(worksheet);

            return isTableFound;
        }

        public void SetVisible(bool check)
        {
            mExcelApp.Visible = check;
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
