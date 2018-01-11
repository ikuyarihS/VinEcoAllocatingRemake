using System;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace VinEcoAllocatingRemake.AllocatingInventory
{
    public class DeleteEvaluationSheetInterop
    {
        /// <summary>
        ///     A special function to deal with Evaluation Sheet created by Aspose.Cells.
        ///     Pirate life ftw.
        /// </summary>
        /// <param name="filePath">Path of file saved and closed by Aspose.Cells.</param>
        [SuppressMessage("ReSharper", "ArgumentsStyleNamedExpression")]
        private void Delete_Evaluation_Sheet_Interop(string filePath)
        {
            try
            {
                // Initialize new instance of Interop Excel.Application.
                var xlApp = new Application
                {
                    ScreenUpdating = false,
                    EnableEvents = false,
                    DisplayAlerts = false,
                    DisplayStatusBar = false,
                    AskToUpdateLinks = false,
                    Visible = false
                };

                Workbooks xlWbs = xlApp.Workbooks;

                // This is hilarious.
                string falseStr = false.ToString();
                string trueStr = true.ToString();

                Workbook xlWb = xlWbs.Open(
                    Filename: filePath,
                    UpdateLinks: falseStr,
                    ReadOnly: falseStr,
                    IgnoreReadOnlyRecommended: trueStr,
                    Origin: XlPlatform.xlWindows,
                    Notify: falseStr,
                    Converter: "0");

                xlApp.Calculation = XlCalculation.xlCalculationManual;

                Sheets xlWss = xlWb.Worksheets;

                //foreach (ExcelInterop.Worksheet worksheet in xlWb.Worksheets)
                for (var sheetIndex = 1; sheetIndex <= xlWss.Count; sheetIndex++)
                {
                    dynamic worksheet = xlWss[sheetIndex.ToString()];
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

                xlWss["1"].Activate();

                Marshal.ReleaseComObject(xlWss);
                xlWb.Close(trueStr);
                Marshal.ReleaseComObject(xlWb);
                Marshal.ReleaseComObject(xlWbs);
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                throw;
            }
        }
    }
}