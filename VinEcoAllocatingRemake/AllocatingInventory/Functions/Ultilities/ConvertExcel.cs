#region

using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

#endregion

namespace VinEcoAllocatingRemake.AllocatingInventory
{
    /// <summary>
    ///     The utilities.
    /// </summary>
    public partial class Utilities
    {
        /// <summary>
        ///     Convert from one file format to another, using Interop.
        ///     Because apparently OpenXML doesn't deal with .xls type ( Including, but not exclusive to .xlsb )
        /// </summary>
        /// <param name="filePath">
        ///     The file Path.
        /// </param>
        /// <param name="previousExtension">
        ///     The previous Extension.
        /// </param>
        /// <param name="afterwardExtension">
        ///     The afterward Extension.
        /// </param>
        /// <param name="yesDeleteFile">
        ///     The yes Delete File.
        /// </param>
        public void ConvertExcelTypeInterop(
            string filePath,
            string previousExtension = "",
            string afterwardExtension = "",
            bool yesDeleteFile = true)
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
                    AskToUpdateLinks = false
                };

                Workbooks workbooks = excelApp.Workbooks;
                Workbook workbook = workbooks.Open(filePath);

                object missing = Type.Missing;

                // This is hilarious.
                string falseStr = false.ToString();

                workbook.SaveAs(
                    filePath.Replace(previousExtension, afterwardExtension),
                    XlFileFormat.xlExcel12,
                    missing,
                    missing,
                    falseStr,
                    falseStr,
                    XlSaveAsAccessMode.xlExclusive,
                    missing,
                    missing,
                    missing);

                workbook.Close(falseStr);
                Release(workbook);

                workbooks.Close();
                Release(workbooks);

                excelApp.Quit();
                Release(excelApp);

                void Release(object suspect)
                {
                    Marshal.ReleaseComObject(suspect);
                    // ReSharper disable once RedundantAssignment
                    suspect = null;
                }

                if (yesDeleteFile) File.Delete(filePath);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                throw;
            }
        }
    }
}