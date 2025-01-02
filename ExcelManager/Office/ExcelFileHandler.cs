using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Data.Common;
using Microsoft.Office.Interop.Excel;

namespace Duck.OfficeAutomationModule.Office
{
    public class ExcelFileHandler : IDisposable
    {
        private Application mExcelApp;
        private Workbook mWorkbook;

        private static readonly string[] MS_TABLE_HEAD = { "126", "127", "128", "129", "130", "131" };
        private const string MS_TABLE_NAME = "name";

        public ExcelFileHandler(string filePath)
        {
            mExcelApp = new Application();
            mWorkbook = mExcelApp.Workbooks.Open(filePath);
        }

        public Sheets GetSheets()
        {
            return mWorkbook.Sheets;
        }

        public void SetVisible(bool check)
        {
            mExcelApp.Visible = check;
        }

        #region MsCetsa
        public bool MsCetsaRun(int sheetIndex, decimal extractionPercentage)
        {
            Debug.Assert(sheetIndex > 0);

            Worksheet experimentWorksheet = (Worksheet)mWorkbook.Sheets[sheetIndex];
            // 가장 작은 인덱스에 있는 값이 있는 셀에서 가장 큰 인덱스에 있는 값이 있는 셀까지가 밤위(빈칸 포함)
            Excel.Range usedRange = experimentWorksheet.UsedRange;

            bool isTableFound = false;
            
            string sourceTableStartAddress = string.Empty;
            string sourceTableEndAddress = string.Empty;
            Excel.Range sourceTableStartCell = null;
            Excel.Range sourceTableEndCell = null;
            for (int i = 1; i <= usedRange.Rows.Count; ++i)
            {
                int checkCount = 0;
                for (int j = 1; j <= usedRange.Columns.Count; ++j)
                {
                    Excel.Range cell = (Excel.Range)usedRange.Cells[i, j];
                    if (cell.Text == MS_TABLE_HEAD[checkCount])
                    {
                        if (checkCount == 0)
                        {
                            sourceTableStartCell = (Excel.Range)usedRange.Cells[i, j - 1];
                        }
                        if (checkCount == MS_TABLE_HEAD.Length - 1)
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
            sourceTableStartCell.Value = MS_TABLE_NAME;
            sourceTableEndCell = sourceTableStartCell.End[XlDirection.xlDown].End[XlDirection.xlToRight];
            sourceTableEndCell = sourceTableEndCell.Offset[0, 1]; // 열을 하나 추가
            sourceTableStartAddress = sourceTableStartCell.Address;
            sourceTableEndAddress = sourceTableEndCell.Address;

            Excel.Range sourceTableRange = experimentWorksheet.Range[sourceTableStartAddress + ":" + sourceTableEndAddress];

            #region 결과 추출

            for (int i = 0; i < 2; ++i)
            {
                // 새 워크 시트에 필터링한 값 출력
                Worksheet newWorksheet = mWorkbook.Sheets.Add();
                newWorksheet.Move(After: mWorkbook.Sheets[mWorkbook.Sheets.Count]);

                // diff 열 계산
                calcMsDiffColumnInTable(sourceTableRange, sourceTableStartCell, sourceTableEndCell, 3, 2);
                extractionMsValueInNewWorksheet(newWorksheet, experimentWorksheet,
                                                sourceTableRange, sourceTableStartCell, sourceTableEndCell
                                                , i, extractionPercentage);

                Marshal.ReleaseComObject(newWorksheet);
            }

            #endregion

            Marshal.ReleaseComObject(sourceTableStartCell);
            Marshal.ReleaseComObject(sourceTableEndCell);
            Marshal.ReleaseComObject(sourceTableRange);
        IS_NOT_FIND_TABLE:
            Marshal.ReleaseComObject(usedRange);
            Marshal.ReleaseComObject(experimentWorksheet);

            return isTableFound;
        }

        private void extractionMsValueInNewWorksheet(
            Excel.Worksheet newWorksheet, Excel.Worksheet experimentWorksheet, Excel.Range sourceTableRange, 
            Excel.Range sourceTableStartCell, Excel.Range sourceTableEndCell,
            int colNum, decimal extractionPercentage)
        {
            // diff열 계산
            calcMsDiffColumnInTable(sourceTableRange, sourceTableStartCell, sourceTableEndCell, 3 + colNum, 2);

            newWorksheet.Cells[1, 1].Value = ((int)(extractionPercentage * 100)).ToString() + "%";

            Excel.Range criteriaRange = newWorksheet.Range["B1:B2"];
            newWorksheet.Range["B4"].Value = MS_TABLE_NAME;
            newWorksheet.Range["C4"].Value = MS_TABLE_HEAD[0];
            newWorksheet.Range["D4"].Value = MS_TABLE_HEAD[1 + colNum];
            newWorksheet.Range["E4"].Value = "diff";
            Excel.Range destinationRange = newWorksheet.Range["B4:E4"];
            criteriaRange.Cells[1, 1].Value = "diff";
            criteriaRange.Cells[2, 1].Value =
                ">=" + experimentWorksheet.Evaluate(
                    $"PERCENTILE.INC({experimentWorksheet.Cells[sourceTableStartCell.Row, sourceTableEndCell.Column].Address}:{experimentWorksheet.Cells[sourceTableEndCell.Row, sourceTableEndCell.Column].Address}, {1 - extractionPercentage})");
            sourceTableRange.AdvancedFilter(
                XlFilterAction.xlFilterCopy, criteriaRange, destinationRange, XlYesNoGuess.xlNo);

            // diff열 계산
            calcMsDiffColumnInTable(sourceTableRange, sourceTableStartCell, sourceTableEndCell, 6 + colNum, 5);

            criteriaRange = newWorksheet.Range["J1:J2"];
            newWorksheet.Range["J4"].Value = MS_TABLE_NAME;
            newWorksheet.Range["K4"].Value = MS_TABLE_HEAD[3];
            newWorksheet.Range["L4"].Value = MS_TABLE_HEAD[4 + colNum];
            newWorksheet.Range["M4"].Value = "diff";
            destinationRange = newWorksheet.Range["J4:M4"];
            criteriaRange.Cells[1, 1].Value = "diff";
            criteriaRange.Cells[2, 1].Value =
                ">=" + experimentWorksheet.Evaluate(
                    $"PERCENTILE.INC({experimentWorksheet.Cells[sourceTableStartCell.Row, sourceTableEndCell.Column].Address}:{experimentWorksheet.Cells[sourceTableEndCell.Row, sourceTableEndCell.Column].Address}, {1 - extractionPercentage})");
            sourceTableRange.AdvancedFilter(
                XlFilterAction.xlFilterCopy, criteriaRange, destinationRange, XlYesNoGuess.xlNo);

            // diff 열 삭제
            removeMsDiffColumnInTable(sourceTableRange, sourceTableStartCell, sourceTableEndCell);

            Marshal.ReleaseComObject(destinationRange);
            Marshal.ReleaseComObject(criteriaRange);
        }

        private void calcMsDiffColumnInTable(
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

        private void removeMsDiffColumnInTable(Excel.Range table, Excel.Range startCell, Excel.Range endCell)
        {
            for (int i = 1; i <= endCell.Row - startCell.Row + 1; i++)
            {
                table.Cells[i, endCell.Column].Value = string.Empty;
            }
        }
        #endregion

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
            catch (Exception e)
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
