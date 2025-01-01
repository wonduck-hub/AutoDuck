using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using OfficeFileHandler;
using System.Diagnostics;
using System.Data.Common;
using Microsoft.Office.Interop.Excel;

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

        private void addDiffColumnInTable(
            Excel.Range table, Excel.Range startCell, Excel.Range endCell, int column1, int column2)
        {
            Debug.Assert(table != null);
            Debug.Assert(startCell != null);
            Debug.Assert(endCell != null);
            Debug.Assert(column1 > 0);
            Debug.Assert(column2 > 0);

            Excel.Range cell1 = null;
            Excel.Range cell2 = null;
            Excel.Range resultCell = null;
            table.Cells[1, endCell.Column].Value = "diff";
            for (int i = 2; i <= endCell.Row - startCell.Row + 1; i++) 
            { 
                cell1 = table.Cells[i, column1];
                cell2 = table.Cells[i, column2]; 
                resultCell = table.Cells[i, endCell.Column]; 
                if (cell1.Value != null && cell2.Value != null) 
                { 
                    resultCell.Value = cell1.Value - cell2.Value; 
                } 
            }

            Marshal.ReleaseComObject(cell1);
            Marshal.ReleaseComObject(cell2);
            Marshal.ReleaseComObject(resultCell);
        }

        private void removeDiffColumnInTable(Excel.Range table, Excel.Range startCell, Excel.Range endCell)
        {
            for (int i = 1; i <= endCell.Row - startCell.Row + 1; i++)
            {
                table.Cells[i, endCell.Column].Value = string.Empty;
            }
        }

        public bool MsCetsaRun(int sheetIndex, decimal extractionPercentage)
        {
            Debug.Assert(sheetIndex > 0);

            Excel.Worksheet experimentWorksheet = (Excel.Worksheet)mWorkbook.Sheets[sheetIndex];
            // 가장 작은 인덱스에 있는 값이 있는 셀에서 가장 큰 인덱스에 있는 값이 있는 셀까지가 밤위(빈칸 포함)
            Excel.Range usedRange = experimentWorksheet.UsedRange;

            bool isTableFound = false;
            const string startHeadName = "name";
            string[] tableHead = { "126", "127", "128", "129", "130", "131" };
            string startTableAddress = string.Empty;
            string endTableAddress = string.Empty;
            Excel.Range startTableCell = null;
            Excel.Range endTableCell = null;
            for (int i = 1; i <= usedRange.Rows.Count; ++i)
            {
                int checkCount = 0;
                for (int j = 1; j <= usedRange.Columns.Count; ++j)
                {
                    Excel.Range cell = (Excel.Range)usedRange.Cells[i, j];
                    if (cell.Text == tableHead[checkCount])
                    {
                        if (checkCount == 0)
                        {
                            startTableCell = ((Excel.Range)usedRange.Cells[i, j - 1]);
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
            // 테이블의 끝 주소 찾기
            startTableCell.Value = startHeadName;
            endTableCell = startTableCell.End[Excel.XlDirection.xlDown].End[Excel.XlDirection.xlToRight];
            endTableCell = endTableCell.Offset[0, 1]; // 열을 하나 추가
            startTableAddress = startTableCell.Address;
            endTableAddress = endTableCell.Address;
            
            Excel.Range tableRange = experimentWorksheet.Range[startTableAddress + ":" + endTableAddress];

            // 새 워크 시트에 필터링한 값 출력
            Excel.Worksheet newSheet = mWorkbook.Sheets.Add();
            newSheet.Move(After: mWorkbook.Sheets[mWorkbook.Sheets.Count]);
            // diff 열 추가
            addDiffColumnInTable(tableRange, startTableCell, endTableCell, 3, 2);
            Excel.Range criteriaRange = newSheet.Range["B1:B2"];
            newSheet.Range["B4"].Value = startHeadName;
            newSheet.Range["C4"].Value = tableHead[0];
            newSheet.Range["D4"].Value = tableHead[1];
            newSheet.Range["E4"].Value = "diff";
            Excel.Range destinationRange = newSheet.Range["B4:E4"];
            criteriaRange.Cells[1, 1].Value = "diff";
            criteriaRange.Cells[2, 1].Value = 
                ">=" + experimentWorksheet.Evaluate(
                    $"PERCENTILE.INC({experimentWorksheet.Cells[startTableCell.Row, endTableCell.Column].Address}:{experimentWorksheet.Cells[endTableCell.Row, endTableCell.Column].Address}, {1 - extractionPercentage})");
            tableRange.AdvancedFilter(
                Excel.XlFilterAction.xlFilterCopy, criteriaRange, destinationRange, Excel.XlYesNoGuess.xlNo);
            // diff 열 삭제
            removeDiffColumnInTable(tableRange, startTableCell, endTableCell);

            // diff 열 추가
            addDiffColumnInTable(tableRange, startTableCell, endTableCell, 6, 5);
            criteriaRange = newSheet.Range["H1:H2"];
            newSheet.Range["H4"].Value = startHeadName;
            newSheet.Range["I4"].Value = tableHead[3];
            newSheet.Range["J4"].Value = tableHead[4];
            newSheet.Range["K4"].Value = "diff";
            destinationRange = newSheet.Range["H4:K4"];
            criteriaRange.Cells[1, 1].Value = "diff";
            criteriaRange.Cells[2, 1].Value =
                ">=" + experimentWorksheet.Evaluate(
                    $"PERCENTILE.INC({experimentWorksheet.Cells[startTableCell.Row, endTableCell.Column].Address}:{experimentWorksheet.Cells[endTableCell.Row, endTableCell.Column].Address}, {1 - extractionPercentage})");
            tableRange.AdvancedFilter(
                Excel.XlFilterAction.xlFilterCopy, criteriaRange, destinationRange, Excel.XlYesNoGuess.xlNo);
            // diff 열 삭제
            removeDiffColumnInTable(tableRange, startTableCell, endTableCell);

            // 새 워크 시트에 필터링한 값 출력
            newSheet = mWorkbook.Sheets.Add();
            newSheet.Move(After: mWorkbook.Sheets[mWorkbook.Sheets.Count]);
            // diff 열 추가
            addDiffColumnInTable(tableRange, startTableCell, endTableCell, 4, 2);
            criteriaRange = newSheet.Range["B1:B2"];
            newSheet.Range["B4"].Value = startHeadName;
            newSheet.Range["C4"].Value = tableHead[0];
            newSheet.Range["D4"].Value = tableHead[2];
            newSheet.Range["E4"].Value = "diff";
            destinationRange = newSheet.Range["B4:E4"];
            criteriaRange.Cells[1, 1].Value = "diff";
            criteriaRange.Cells[2, 1].Value =
                ">=" + experimentWorksheet.Evaluate(
                    $"PERCENTILE.INC({experimentWorksheet.Cells[startTableCell.Row, endTableCell.Column].Address}:{experimentWorksheet.Cells[endTableCell.Row, endTableCell.Column].Address}, {1 - extractionPercentage})");
            tableRange.AdvancedFilter(
                Excel.XlFilterAction.xlFilterCopy, criteriaRange, destinationRange, Excel.XlYesNoGuess.xlNo);
            // diff 열 삭제
            removeDiffColumnInTable(tableRange, startTableCell, endTableCell);

            // diff 열 추가
            addDiffColumnInTable(tableRange, startTableCell, endTableCell, 7, 5);
            criteriaRange = newSheet.Range["H1:H2"];
            newSheet.Range["H4"].Value = startHeadName;
            newSheet.Range["I4"].Value = tableHead[3];
            newSheet.Range["J4"].Value = tableHead[5];
            newSheet.Range["K4"].Value = "diff";
            destinationRange = newSheet.Range["H4:K4"];
            criteriaRange.Cells[1, 1].Value = "diff";
            criteriaRange.Cells[2, 1].Value =
                ">=" + experimentWorksheet.Evaluate(
                    $"PERCENTILE.INC({experimentWorksheet.Cells[startTableCell.Row, endTableCell.Column].Address}:{experimentWorksheet.Cells[endTableCell.Row, endTableCell.Column].Address}, {1 - extractionPercentage})");
            tableRange.AdvancedFilter(
                Excel.XlFilterAction.xlFilterCopy, criteriaRange, destinationRange, Excel.XlYesNoGuess.xlNo);
            // diff 열 삭제
            removeDiffColumnInTable(tableRange, startTableCell, endTableCell);

            Marshal.ReleaseComObject(destinationRange);
            Marshal.ReleaseComObject(criteriaRange);
            Marshal.ReleaseComObject(startTableCell);
            Marshal.ReleaseComObject(endTableCell);
            Marshal.ReleaseComObject(newSheet);
            Marshal.ReleaseComObject(tableRange);
        IS_NOT_FIND_TABLE:
            Marshal.ReleaseComObject(usedRange);
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
