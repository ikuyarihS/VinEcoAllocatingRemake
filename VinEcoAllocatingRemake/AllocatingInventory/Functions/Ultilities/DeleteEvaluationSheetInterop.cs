namespace VinEcoAllocatingRemake.AllocatingInventory
{
    #region

    using System;
    using System.Diagnostics;
    using System.Diagnostics.CodeAnalysis;
    using System.Runtime.InteropServices;

    using Microsoft.Office.Interop.Excel;

    #endregion

    /// <summary>
    ///     The delete evaluation sheet interop.
    /// </summary>
    [SuppressMessage("ReSharper", "ArrangeThisQualifier")]
    public partial class Utilities
    {
        /// <summary>
        ///     A special function to deal with Evaluation Sheet created by Aspose.Cells.
        ///     Pirate life ftw.
        /// </summary>
        /// <param name="filePath"> Path of file saved and closed by Aspose.Cells. </param>
        public void DeleteEvaluationSheetInterop(string filePath)
        {
            try
            {
                // Initialize new instance of Interop Excel.Application.
                var excelApp = new Application
                                   {
                                       ScreenUpdating = false,
                                       EnableEvents = false,
                                       DisplayAlerts = false,
                                       DisplayStatusBar = false,
                                       AskToUpdateLinks = false,
                                       Visible = false
                                   };

                Workbooks workbooks = excelApp.Workbooks;

                // This is hilarious.
                string falseStr = false.ToString();
                string trueStr = true.ToString();

                Workbook workbook = workbooks.Open(
                    filePath,
                    falseStr,
                    falseStr,
                    IgnoreReadOnlyRecommended: trueStr,
                    Origin: XlPlatform.xlWindows,
                    Notify: falseStr,
                    Converter: "0");

                excelApp.Calculation = XlCalculation.xlCalculationManual;

                Sheets worksheets = workbook.Worksheets;

                // foreach (ExcelInterop.Worksheet worksheet in xlWb.Worksheets)
                for (var sheetIndex = 1; sheetIndex <= worksheets.Count; sheetIndex++)
                {
                    dynamic worksheet = worksheets[sheetIndex.ToString()];
                    switch (worksheet.Name)
                    {
                        case "Config":
                            worksheet.Cells[1, 1].Value2 = trueStr;
                            break;
                        case "Evaluation Warning":
                            worksheet.Delete();
                            break;

                        // ReSharper disable once RedundantEmptySwitchSection
                        default:
                            break;
                    }

                    Marshal.ReleaseComObject(worksheet);
                }

                worksheets[this.IntToObject(1)].Activate();

                Marshal.ReleaseComObject(worksheets);
                workbook.Close(trueStr);
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(workbooks);
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                throw;
            }
        }
    }
}