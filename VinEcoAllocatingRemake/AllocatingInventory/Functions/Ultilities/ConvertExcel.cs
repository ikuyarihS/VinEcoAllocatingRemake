namespace VinEcoAllocatingRemake.AllocatingInventory
{
    #region

    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.IO;
    using System.Runtime.InteropServices;

    using Aspose.Cells;

    using Microsoft.Office.Interop.Excel;

    using Workbook = Aspose.Cells.Workbook;

    #endregion

    /// <summary>
    ///     The utilities.
    /// </summary>
    public partial class Utilities
    {
        /// <summary>
        ///     The convert excel type aspose.
        /// </summary>
        /// <param name="filePath"> The file path. </param>
        /// <param name="afterwardExtension"> The afterward extension. </param>
        /// <param name="yesDeleteFile"> The yes delete file. </param>
        public void ConvertExcelTypeAspose(
            string filePath,
            string afterwardExtension,
            bool yesDeleteFile = true)
        {
            try
            {
                // Initialize new instance of Aspose.Cells Workbook
                var workbook = new Workbook(filePath, new LoadOptions { MemorySetting = MemorySetting.MemoryPreference });

                var dicExtension = new Dictionary<string, SaveFormat>
                                       {
                                           { "xlsx", SaveFormat.Xlsx },
                                           { "xlsb", SaveFormat.Xlsb },
                                           { "xlsm", SaveFormat.Xlsm },
                                           { "pdf", SaveFormat.Pdf }
                                       };

                workbook.Save(filePath, dicExtension[afterwardExtension]);

                if (yesDeleteFile)
                {
                    File.Delete(filePath);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                throw;
            }
        }

        /// <summary>
        ///     Convert from one file format to another, using Interop.
        ///     Because apparently OpenXML doesn't deal with .xls type ( Including, but not exclusive to .xlsb )
        /// </summary>
        /// <param name="filePath"> The file Path. </param>
        /// <param name="previousExtension"> The previous Extension. </param>
        /// <param name="afterwardExtension"> The afterward Extension. </param>
        /// <param name="yesDeleteFile"> The yes Delete File. </param>
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
                Microsoft.Office.Interop.Excel.Workbook workbook = workbooks.Open(filePath);

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

                if (yesDeleteFile)
                {
                    File.Delete(filePath);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                throw;
            }
        }
    }
}