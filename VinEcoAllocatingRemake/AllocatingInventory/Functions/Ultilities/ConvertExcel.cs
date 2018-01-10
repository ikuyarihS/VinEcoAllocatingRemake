using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace VinEcoAllocatingRemake.AllocatingInventory
{
    public partial class Utilities
    {
        /// <summary>
        ///     Convert from one file format to another, using Interop.
        ///     Because apparently OpenXML doesn't deal with .xls type ( Including, but not exclusive to .xlsb )
        /// </summary>
        public void ConvertExcelTypeInterop(
            string filePath,
            string previousExtension = "",
            string afterwardExtension = "",
            bool yesDeleteFile = true)
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
                    AskToUpdateLinks = false
                };
                
                Workbook xlWb = xlApp.Workbooks.Open(filePath);

                object missing = Type.Missing;

                xlWb.SaveAs(
                    Filename: filePath.Replace(previousExtension, afterwardExtension), 
                    FileFormat: XlFileFormat.xlExcel12, 
                    Password: missing,
                    WriteResPassword: missing, 
                    ReadOnlyRecommended: false, 
                    CreateBackup: false, 
                    AccessMode: XlSaveAsAccessMode.xlExclusive, 
                    ConflictResolution: missing, 
                    AddToMru: missing, 
                    TextCodepage: missing);

                xlWb.Close(false);
                Marshal.ReleaseComObject(xlWb);

                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);

                if (yesDeleteFile)
                    File.Delete(filePath);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                throw;
            }
        }
    }
}