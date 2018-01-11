using System;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.IO;
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
        [SuppressMessage("ReSharper", "ArgumentsStyleOther")]
        [SuppressMessage("ReSharper", "ArgumentsStyleNamedExpression")]
        [SuppressMessage("ReSharper", "MemberCanBeMadeStatic.Global")]
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

                Workbooks xlWbs = xlApp.Workbooks;
                Workbook xlWb = xlWbs.Open(Filename: filePath);

                object missing = Type.Missing;

                // This is hilarious.
                string falseStr = false.ToString();

                xlWb.SaveAs(
                    Filename: filePath.Replace(oldValue: previousExtension, newValue: afterwardExtension),
                    FileFormat: XlFileFormat.xlExcel12,
                    Password: missing,
                    WriteResPassword: missing,
                    ReadOnlyRecommended: falseStr,
                    CreateBackup: falseStr,
                    AccessMode: XlSaveAsAccessMode.xlExclusive,
                    ConflictResolution: missing,
                    AddToMru: missing,
                    TextCodepage: missing);

                xlWb.Close(SaveChanges: falseStr);
                Release(xlWb);

                xlWbs.Close();
                Release(xlWbs);

                xlApp.Quit();
                Release(xlApp);

                void Release(object suspect)
                {
                    Marshal.ReleaseComObject(suspect);
                    suspect = null;
                }

                if (yesDeleteFile) File.Delete(path: filePath);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(message: ex.Message);
                throw;
            }
        }
    }
}