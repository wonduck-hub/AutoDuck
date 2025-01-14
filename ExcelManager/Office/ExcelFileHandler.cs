using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Data.Common;

namespace Duck.OfficeAutomationModule.Office
{
    public class ExcelFileHandler : IDisposable
    {
        private Excel.Application mExcelApp;
        private Excel.Workbook mWorkbook;

        private static readonly string[] MS_TABLE_HEAD = { "126", "127", "128", "129", "130", "131" };
        private const string MS_TABLE_NAME = "name";

        static public bool IsExcelInstalled()
        {
            try { 
                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null) 
                { 
                    return false; 
                } 
                dynamic excelApp = Activator.CreateInstance(excelType); 
                return true; 
            }
            catch 
            { 
                return false;
            }
        }

        public ExcelFileHandler(string filePath)
        {
            mExcelApp = new Excel.Application();
            mWorkbook = mExcelApp.Workbooks.Open(filePath);
        }

        public Excel.Sheets GetSheets()
        {
            return mWorkbook.Sheets;
        }

        public void SetVisible(bool check)
        {
            mExcelApp.Visible = check;
        }

        #region CETSA-MS
        public async Task<bool> CetsaMsRun(int sheetIndex, decimal extractionPercentage)
        {
            Debug.Assert(sheetIndex > 0);

            Excel.Worksheet experimentWorksheet = (Excel.Worksheet)mWorkbook.Sheets[sheetIndex];
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
                }
            }

            if (!isTableFound)
            {
                goto IS_NOT_FIND_TABLE;
            }

        IS_FIND_TABLE:
            // 테이블의 끝 주소 찾기
            sourceTableStartCell.Value = MS_TABLE_NAME;
            sourceTableEndCell = sourceTableStartCell.End[Excel.XlDirection.xlDown].End[Excel.XlDirection.xlToRight];
            sourceTableEndCell = sourceTableEndCell.Offset[0, 1]; // 열을 하나 추가해 diff 열로 사용
            sourceTableStartAddress = sourceTableStartCell.Address;
            sourceTableEndAddress = sourceTableEndCell.Address;

            Excel.Range sourceTableRange = experimentWorksheet.Range[sourceTableStartAddress + ":" + sourceTableEndAddress];

            #region 결과 추출

            for (int i = 0; i < 2; ++i)
            {
                // 새 워크 시트에 필터링한 값 출력
                Excel.Worksheet newWorksheet = mWorkbook.Sheets.Add();
                newWorksheet.Move(After: mWorkbook.Sheets[mWorkbook.Sheets.Count]);

                // diff 열 계산
                await extractionMsValueInNewWorksheet(newWorksheet, experimentWorksheet,
                                                sourceTableRange, sourceTableStartCell, sourceTableEndCell
                                                , i, extractionPercentage);

                Marshal.ReleaseComObject(newWorksheet);
            }

            #endregion

        IS_NOT_FIND_TABLE:
            return isTableFound;
        }

        private async Task extractionMsValueInNewWorksheet(
            Excel.Worksheet newWorksheet, Excel.Worksheet experimentWorksheet, Excel.Range sourceTableRange, 
            Excel.Range sourceTableStartCell, Excel.Range sourceTableEndCell,
            int colNum, decimal extractionPercentage)
        {
            // TODO: 여기서 추출한 단백질에 대한 정보 출력 코드 추가
            // diff열 계산
            calcMsDiffColumnInTable(experimentWorksheet, sourceTableStartCell, sourceTableEndCell, 3 + colNum, 2);

            newWorksheet.Cells[1, 1].Value = ((int)(extractionPercentage * 100)).ToString() + "%";

            Excel.Range criteriaRange = newWorksheet.Range["B1:B2"];
            newWorksheet.Range["B4"].Value = MS_TABLE_NAME;
            newWorksheet.Range["C4"].Value = MS_TABLE_HEAD[0];
            newWorksheet.Range["D4"].Value = MS_TABLE_HEAD[1 + colNum];
            newWorksheet.Range["E4"].Value = "diff";
            Excel.Range destinationRange = newWorksheet.Range["B4:E4"];
            criteriaRange.Cells[1, 1].Value = "diff";
            criteriaRange.Cells[2, 1].Value = 
                ">=" + experimentWorksheet.Evaluate($"Log({1 + extractionPercentage}, 2)");
            sourceTableRange.AdvancedFilter(
                Excel.XlFilterAction.xlFilterCopy, criteriaRange, destinationRange, Excel.XlYesNoGuess.xlNo);

            // diff열 계산
            calcMsDiffColumnInTable(experimentWorksheet, sourceTableStartCell, sourceTableEndCell, 6 + colNum, 5);

            criteriaRange = newWorksheet.Range["J1:J2"];
            newWorksheet.Range["J4"].Value = MS_TABLE_NAME;
            newWorksheet.Range["K4"].Value = MS_TABLE_HEAD[3];
            newWorksheet.Range["L4"].Value = MS_TABLE_HEAD[4 + colNum];
            newWorksheet.Range["M4"].Value = "diff";
            destinationRange = newWorksheet.Range["J4:M4"];
            criteriaRange.Cells[1, 1].Value = "diff";
            criteriaRange.Cells[2, 1].Value =
                ">=" + experimentWorksheet.Evaluate($"Log({1 + extractionPercentage}, 2)");
            sourceTableRange.AdvancedFilter(
                Excel.XlFilterAction.xlFilterCopy, criteriaRange, destinationRange, Excel.XlYesNoGuess.xlNo);

            // diff 열 삭제
            removeMsDiffColumnInTable(experimentWorksheet, sourceTableStartCell, sourceTableEndCell);
        }

        #region calcDiff
        private void calcMsDiffColumnInTable(
            Excel.Worksheet experimentWorksheet, Excel.Range startCell, Excel.Range endCell, int column1, int column2)
        {
            Debug.Assert(experimentWorksheet != null);
            Debug.Assert(startCell != null);
            Debug.Assert(endCell != null);
            Debug.Assert(column1 > 0);
            Debug.Assert(column2 > 0);

            string formulaString = 
                "=" + ((Excel.Range)experimentWorksheet.Cells[startCell.Row + 1, column1]).Address[false, false] + 
                "-" + ((Excel.Range)experimentWorksheet.Cells[startCell.Row + 1, column2]).Address[false, false];

            Excel.Range sourceRange = experimentWorksheet.Cells[startCell.Row + 1, endCell.Column];
            experimentWorksheet.Cells[startCell.Row, endCell.Column].Value = "diff";
            Excel.Range destinationRange =
                experimentWorksheet.Range[experimentWorksheet.Cells[startCell.Row + 1, endCell.Column],
                                          experimentWorksheet.Cells[endCell.Row, endCell.Column]];

            sourceRange.Formula = formulaString;
            sourceRange.AutoFill(destinationRange, Excel.XlAutoFillType.xlFillCopy);
        }

        private void removeMsDiffColumnInTable(
            Excel.Worksheet experimentWorksheet, Excel.Range startCell, Excel.Range endCell)
        {
            Debug.Assert(experimentWorksheet != null);
            Debug.Assert(startCell != null);
            Debug.Assert(endCell != null);

            experimentWorksheet.Cells[startCell.Row, endCell.Column].Value = string.Empty;
            Excel.Range destinationRange =
                experimentWorksheet.Range[experimentWorksheet.Cells[startCell.Row, endCell.Column],
                                          experimentWorksheet.Cells[endCell.Row, endCell.Column]];

            experimentWorksheet.Cells[startCell.Row, endCell.Column].AutoFill(
                destinationRange, Excel.XlAutoFillType.xlFillCopy);
        }
        #endregion
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
