using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace VinEcoAllocatingRemake.AllocatingInventory
{
    public class DeleteEvaluationSheetInterop
    {
        /// <summary>
        ///     A special function to deal with Evaluation Sheet created by Aspose.Cells.
        ///     Pirate life ftw.
        /// </summary>
        /// <param name="filePath">Path of file saved and closed by Aspose.Cells.</param>
        private void Delete_Evaluation_Sheet_Interop(string filePath)
        {
            try
            {
                // Initialize new instance of Interop Excel.Application.
                var xlApp = new Microsoft.Office.Interop.Excel.Application
                {
                    ScreenUpdating = false,
                    EnableEvents = false,
                    DisplayAlerts = false,
                    DisplayStatusBar = false,
                    AskToUpdateLinks = false,
                    Visible = false
                };

                Microsoft.Office.Interop.Excel.Workbooks xlWbs = xlApp.Workbooks;

                //ExcelInterop.Workbook xlWb = xlWbs.Open(
                //    Filename: filePath,
                //    UpdateLinks: false,
                //    ReadOnly: false,
                //    Format: 5,
                //    Password: string.Empty,
                //    WriteResPassword: string.Empty,
                //    IgnoreReadOnlyRecommended: true,
                //    Origin: ExcelInterop.XlPlatform.xlWindows,
                //    Delimiter: string.Empty,
                //    Editable: true,
                //    Notify: false,
                //    Converter: 0,
                //    AddToMru: true,
                //    Local: false,
                //    CorruptLoad: false);
                
                // This is hilarious.
                string falseStr = false.ToString();
                string trueStr = true.ToString();

                Microsoft.Office.Interop.Excel.Workbook xlWb = xlWbs.Open(
                    Filename: filePath,
                    UpdateLinks: falseStr,
                    ReadOnly: falseStr,
                    IgnoreReadOnlyRecommended: trueStr,
                    Origin: Microsoft.Office.Interop.Excel.XlPlatform.xlWindows,
                    Notify: falseStr,
                    Converter: "0");

                xlApp.Calculation = Microsoft.Office.Interop.Excel.XlCalculation.xlCalculationManual;

                Microsoft.Office.Interop.Excel.Sheets xlWss = xlWb.Worksheets;

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
                xlWb.Close(SaveChanges: trueStr);
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