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

        private void createResultWorksheet(
            Excel.Range experimentTableRange, int tableStartRow, int tableStartColumn, 
            int comparison1ColumnIndex, int chemicalSubstanceResult1ColumnIndex,
            int comparison2ColumnIndex, int chemicalSubstanceResult2ColumnIndex)
        {
            Excel.Worksheet resultSheet = mWorkbook.Sheets.Add();
            resultSheet.Move(After: mWorkbook.Sheets[mWorkbook.Sheets.Count]);
            // 열 제목
            resultSheet.Cells[2, 2] =
                    experimentTableRange.Cells[tableStartRow, chemicalSubstanceResult1ColumnIndex];
            resultSheet.Cells[2, 3] =
                    experimentTableRange.Cells[tableStartRow, chemicalSubstanceResult2ColumnIndex];

            // 추출
            int rowIndex = 1;
            int rowMerge = 2;
            while (((Excel.Range)experimentTableRange.Cells[tableStartRow + rowIndex, comparison1ColumnIndex]).Text != string.Empty)
            {
                resultSheet.Cells[rowIndex + rowMerge, 1] =
                    experimentTableRange.Cells[tableStartRow + rowIndex, tableStartColumn];

                double result1 = 0;
                double result2 = 0;
                if (double.TryParse(experimentTableRange.Cells[tableStartRow + rowIndex, chemicalSubstanceResult1ColumnIndex].Text, out result1)
                    && double.TryParse(experimentTableRange.Cells[tableStartRow + rowIndex, comparison1ColumnIndex].Text, out result2))
                {
                    resultSheet.Cells[rowIndex + rowMerge, 2].Value = result1 - result2;
                }
                else
                {
                    resultSheet.Cells[rowIndex + rowMerge, 2].Value = "error";
                }

                if (double.TryParse(experimentTableRange.Cells[tableStartRow + rowIndex, chemicalSubstanceResult2ColumnIndex].Text, out result1)
                    && double.TryParse(experimentTableRange.Cells[tableStartRow + rowIndex, comparison2ColumnIndex].Text, out result2))
                {
                    resultSheet.Cells[rowIndex + rowMerge, 3].Value = result1 - result2;
                }
                else
                {
                    resultSheet.Cells[rowIndex + rowMerge, 3].Value = "error";
                }

                ++rowIndex;
            }

            Marshal.ReleaseComObject(resultSheet);
        }

        public bool MsCetsaRun(int sheetIndex)
        {
            Debug.Assert(sheetIndex > 0);

            Excel.Worksheet experimentWorksheet = (Excel.Worksheet)mWorkbook.Sheets[sheetIndex];
            // 가장 작은 인덱스에 있는 값이 있는 셀에서 가장 큰 인덱스에 있는 값이 있는 셀까지가 밤위(빈칸 포함)
            Excel.Range experimentTableRange = experimentWorksheet.UsedRange;

            bool isTableFound = false;
            string[] tableHead = { "126", "127", "128", "129", "130", "131" };
            int tableStartRow = 0;
            int tableStartColumn = 0;
            for (int i = 1; i <= experimentTableRange.Rows.Count; ++i)
            {
                int checkCount = 0;
                for (int j = 1; j <= experimentTableRange.Columns.Count; ++j)
                {
                    Excel.Range cell = (Excel.Range)experimentTableRange.Cells[i, j];
                    if (cell.Text == tableHead[checkCount])
                    {
                        if (checkCount == 0)
                        {
                            tableStartRow = i;
                            tableStartColumn = j - 1;
                        }
                        if (checkCount == tableHead.Length - 1)
                        {
                            isTableFound = true;
                            goto IS_FIND_TABLE;
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
                goto IS_NOT_FIND_TABLE;
            }

        IS_FIND_TABLE:

            createResultWorksheet(experimentTableRange, tableStartRow, tableStartColumn, 
                tableStartColumn + 1, tableStartColumn + 2, tableStartColumn + 4, tableStartColumn + 5);
            createResultWorksheet(experimentTableRange, tableStartRow, tableStartColumn,
                tableStartColumn + 1, tableStartColumn + 3, tableStartColumn + 4, tableStartColumn + 6);

        IS_NOT_FIND_TABLE:
            Marshal.ReleaseComObject(experimentTableRange);
            Marshal.ReleaseComObject(experimentWorksheet);

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
