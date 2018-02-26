// --------------------------------------------------------------------------------------------------------------------
// <copyright file="DeleteEvaluationSheetInterop.cs" company="VinEco">
//   Shirayuki 2018.
// </copyright>
// <summary>
//   The delete evaluation sheet interop.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

using System;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace VinEcoAllocatingRemake.AllocatingInventory
{
    #region

    #endregion

    /// <summary>
    ///     The delete evaluation sheet interop.
    /// </summary>
    // ReSharper disable once StyleCop.SA1404
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
                                       ScreenUpdating   = false,
                                       EnableEvents     = false,
                                       DisplayAlerts    = false,
                                       DisplayStatusBar = false,
                                       AskToUpdateLinks = false,
                                       Visible          = false
                                   };

                Workbooks workbooks = excelApp.Workbooks;

                Workbook workbook = workbooks.Open(
                    filePath,
                    false,
                    false,
                    IgnoreReadOnlyRecommended: true,
                    Origin: XlPlatform.xlWindows,
                    Notify: false,
                    Converter: 0);

                excelApp.Calculation = XlCalculation.xlCalculationManual;

                Sheets worksheets = workbook.Worksheets;

                // foreach (ExcelInterop.Worksheet worksheet in xlWb.Worksheets)
                for (var sheetIndex = 1; sheetIndex <= worksheets.Count; sheetIndex++)
                {
                    Worksheet worksheet = worksheets[sheetIndex];

                    if (worksheet.Name == "Config")
                    {
                        worksheet.Cells[1, 1].Value2 = true;
                    }

                    if (worksheet.Name.IndexOf("Evaluation Warning", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        worksheet.Delete();
                    }

                    Marshal.ReleaseComObject(worksheet);
                }

                worksheets[1].Activate();

                Marshal.ReleaseComObject(worksheets);

                workbook.Save();
                workbook.Close();
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