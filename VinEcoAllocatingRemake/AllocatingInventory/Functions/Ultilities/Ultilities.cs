using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Bold = DocumentFormat.OpenXml.Spreadsheet.Bold;
using Border = DocumentFormat.OpenXml.Spreadsheet.Border;
using ExcelInterop = Microsoft.Office.Interop.Excel; // I fucking hate this.
using Fonts = DocumentFormat.OpenXml.Spreadsheet.Fonts;

// ReSharper disable ArgumentsStyleNamedExpression
// ReSharper disable ArgumentsStyleLiteral
// ReSharper disable ArgumentsStyleStringLiteral

namespace VinEcoAllocatingRemake.AllocatingInventory
{
    public partial class AllocatingInventory
    {
        private string ReturnOnlyInitial(string stringKey, Dictionary<string, string> stringsProcessed)
        {
            try
            {
                lock (stringsProcessed)
                {
                    if (stringsProcessed.TryGetValue(stringKey, out string stringResult))
                        return stringResult;

                    stringResult = string.Join(string.Empty, stringKey.Split(' ').Select(x => x.First()));
                    stringsProcessed.Add(stringKey, stringResult);
                    return stringResult;
                }
            }
            catch (Exception ex)
            {
                WriteToRichTextBoxOutput(ex.Message);
                throw;
            }
        }

        /// <summary>
        ///     A special function to deal with Evaluation Sheet created by Aspose.Cells.
        ///     Pirate life ftw.
        /// </summary>
        /// <param name="filePath">Path of file saved and closed by Aspose.Cells.</param>
        private void Delete_Evaluation_Sheet_Interop(string filePath)
        {
            try
            {
                // Remember the list of running Excel.Application.
                // Before initialize xlApp.
                //var processBefore = Process.GetProcessesByName("excel");

                // Initialize new instance of Interop Excel.Application.
                var xlApp = new ExcelInterop.Application
                {
                    ScreenUpdating = false,
                    EnableEvents = false,
                    DisplayAlerts = false,
                    DisplayStatusBar = false,
                    AskToUpdateLinks = false,
                    Visible = false
                };

                // Remember the list of running Excel.Application.
                // After initialize xlApp.
                //var processAfter = Process.GetProcessesByName("excel");

                //var processId = processAfter.Except(processBefore).Last().Id;

                ExcelInterop.Workbooks xlWbs = xlApp.Workbooks;

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

                ExcelInterop.Workbook xlWb = xlWbs.Open(
                    Filename: filePath,
                    UpdateLinks: false,
                    ReadOnly: false,
                    IgnoreReadOnlyRecommended: true,
                    Origin: ExcelInterop.XlPlatform.xlWindows,
                    Notify: false,
                    Converter: 0);

                xlApp.Calculation = ExcelInterop.XlCalculation.xlCalculationManual;

                ExcelInterop.Sheets xlWss = xlWb.Worksheets;

                //foreach (ExcelInterop.Worksheet worksheet in xlWb.Worksheets)
                for (var sheetIndex = 1; sheetIndex <= xlWss.Count; sheetIndex++)
                {
                    dynamic worksheet = xlWss[sheetIndex];
                    switch (worksheet.Name)
                    {
                        case "Config":
                            worksheet.Cells[1, 1].Value2 = true;
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

                xlWss[1].Activate();

                Marshal.ReleaseComObject(xlWss);
                xlWb.Close(SaveChanges: true);
                Marshal.ReleaseComObject(xlWb);
                Marshal.ReleaseComObject(xlWbs);
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
            }
            catch (Exception ex)
            {
                WriteToRichTextBoxOutput(ex.Message);
                throw;
            }
        }

        /// <summary>
        ///     Appending desired text into Output RichTextBox.
        ///     By default, it will be in a new line.
        /// </summary>
        /// <param name="message">Message for RichTextBox to Append.</param>
        /// <param name="importanceLevel">
        ///     Level of Importance.
        ///     0 = Default.
        ///     1 = Very importanto.
        ///     2 = Meh.
        /// </param>
        /// <param name="newLine">A seperated new line?</param>
        /// <param name="hasTimeStamp">Include Time Stamp</param>
        private void WriteToRichTextBoxOutput(object message = null, byte importanceLevel = 0, bool newLine = true,
            bool hasTimeStamp = true)
        {
            void Action()
            {
                try
                {
                    var textRange = new TextRange(RichTextBoxOutput.Document.ContentEnd,
                        RichTextBoxOutput.Document.ContentEnd);
                    var brushConverter = new BrushConverter();

                    string extraMessage = string.Empty;
                    if (message == null || message.ToString() == string.Empty) message = string.Empty;
                    switch (importanceLevel)
                    {
                        case 0:
                            break;
                        case 1:
                            extraMessage = "!!! - ";
                            break;
                        case 2:
                            extraMessage = "      ";
                            break;
                        default:
                            extraMessage = string.Empty;
                            break;
                    }

                    if (hasTimeStamp && message != null && message.ToString() != string.Empty)
                        ExtraTimeStamp();

                    textRange.Text =
                        $"{(hasTimeStamp ? extraMessage : string.Empty)}{message}{(newLine ? "\r" : " ")}";
                    textRange.ApplyPropertyValue(TextElement.ForegroundProperty,
                        brushConverter.ConvertFromString("Cornflowerblue") ??
                        throw new InvalidOperationException("What the heck?"));
                }

                catch (Exception ex)
                {
                    WriteToRichTextBoxOutput(ex.Message);
                    throw;
                }
            }

            Application.Current.Dispatcher.BeginInvoke((Action) Action);
        }

        private void RichTextBoxOutput_TextChanged(object sender, TextChangedEventArgs e)
        {
            RichTextBoxOutput.ScrollToEnd();
        }

        private void ExtraTimeStamp()
        {
            try
            {
                var textRange = new TextRange(RichTextBoxOutput.Document.ContentEnd,
                    RichTextBoxOutput.Document.ContentEnd) {Text = DateTime.Now.ToString("[HH:mm] ")};

                /* $"[{DateTime.Now:HH:mm}] ";*/

                textRange.ApplyPropertyValue(TextElement.ForegroundProperty,
                    new BrushConverter().ConvertFromString("ForestGreen") ??
                    throw new InvalidOperationException("What the heck?"));
            }

            catch (Exception ex)
            {
                WriteToRichTextBoxOutput(ex.Message);
                throw;
            }
        }

        /// <summary>
        ///     A simple helper to get the column name from the column index. This is not well tested!
        ///     <para />
        ///     Worked anyway. For a Dictionary anyway.
        /// </summary>
        private static string GetColumnName(int columnIndex)
        {
            int dividend = columnIndex;
            string columnName = string.Empty;

            while (dividend > 0)
            {
                int modifier = (dividend - 1) % 26;
                columnName = $"{Convert.ToChar(65 + modifier)}{columnName}";
                dividend = (dividend - modifier) / 26;
            }

            return columnName;
        }


        /// <summary>
        ///     OpenWriter Style, for Multiple DataTable into Multiple Worksheets in a single Workbook. A real fucking pain.
        /// </summary>
        public static void LargeExportOneWorkbook(string filePath, List<DataTable> listDt, bool yesNoHeader = false,
            bool yesNoZero = false)
        {
            try
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
                {
                    document.AddWorkbookPart();

                    OpenXmlWriter writerXb = OpenXmlWriter.Create(document.WorkbookPart);
                    writerXb.WriteStartElement(new Workbook());
                    writerXb.WriteStartElement(new Sheets());

                    var count = 0;

                    foreach (DataTable dt in listDt)
                    {
                        var dicType = new Dictionary<Type, CellValues>();

                        var dicColName = new Dictionary<int, string>();

                        for (var colIndex = 0; colIndex < dt.Columns.Count; colIndex++)
                            dicColName.Add(colIndex + 1, GetColumnName(colIndex + 1));

                        dicType.Add(typeof(DateTime), CellValues.Date);
                        dicType.Add(typeof(string), CellValues.InlineString);
                        dicType.Add(typeof(double), CellValues.Number);
                        dicType.Add(typeof(int), CellValues.Number);

                        //this list of attributes will be used when writing a start element
                        List<OpenXmlAttribute> attributes;

                        var workSheetPart = document.WorkbookPart.AddNewPart<WorksheetPart>();

                        OpenXmlWriter writer = OpenXmlWriter.Create(workSheetPart);
                        writer.WriteStartElement(new Worksheet());
                        writer.WriteStartElement(new SheetData());

                        if (yesNoHeader)
                        {
                            //create a new list of attributes
                            attributes = new List<OpenXmlAttribute>();
                            // add the row index attribute to the list
                            attributes.Add(new OpenXmlAttribute("r", null, 1.ToString()));

                            //write the row start element with the row index attribute
                            writer.WriteStartElement(new Row(), attributes);

                            for (var columnNum = 1; columnNum <= dt.Columns.Count; ++columnNum)
                            {
                                //reset the list of attributes
                                attributes = new List<OpenXmlAttribute>
                                {
                                    new OpenXmlAttribute("t", null, "str"),
                                    new OpenXmlAttribute("r", "", $"{dicColName[columnNum]}1")
                                };
                                // add data type attribute - in this case inline string (you might want to look at the shared strings table)
                                //add the cell reference attribute

                                //write the cell start element with the type and reference attributes
                                writer.WriteStartElement(new Cell(), attributes);

                                //write the cell value
                                writer.WriteElement(new CellValue(dt.Columns[columnNum - 1].ColumnName));

                                //writer.WriteElement(new CellValue(string.Format("This is Row {0}, Cell {1}", rowNum, columnNum)));

                                // write the end cell element
                                writer.WriteEndElement();
                            }

                            // write the end row element
                            writer.WriteEndElement();
                        }

                        for (var rowNum = 1; rowNum <= dt.Rows.Count; rowNum++)
                        {
                            //create a new list of attributes
                            attributes = new List<OpenXmlAttribute>
                            {
                                new OpenXmlAttribute("r", null,
                                    (yesNoHeader ? rowNum + 1 : rowNum).ToString())
                            };
                            // add the row index attribute to the list

                            //write the row start element with the row index attribute
                            writer.WriteStartElement(new Row(), attributes);

                            DataRow dr = dt.Rows[rowNum - 1];
                            for (var columnNum = 1; columnNum <= dt.Columns.Count; columnNum++)
                            {
                                Type type = dt.Columns[columnNum - 1].DataType;
                                //reset the list of attributes
                                attributes = new List<OpenXmlAttribute>
                                {
                                    new OpenXmlAttribute("t", null,
                                        type == typeof(string) ? "str" : dicType[type].ToString()),
                                    new OpenXmlAttribute("r", "",
                                        $"{dicColName[columnNum]}{(yesNoHeader ? rowNum + 1 : rowNum)}")
                                };
                                // add data type attribute - in this case inline string (you might want to look at the shared strings table)
                                //add the cell reference attribute

                                //write the cell start element with the type and reference attributes
                                writer.WriteStartElement(new Cell(), attributes);

                                //write the cell value
                                if (yesNoZero | (dr[columnNum - 1].ToString() != "0"))
                                    writer.WriteElement(new CellValue(dr[columnNum - 1].ToString()));
                                {
                                    // In case of 0. Can safely forsake this part.
                                    //writer.WriteElement(new CellValue(""));
                                }

                                //writer.WriteElement(new CellValue(string.Format("This is Row {0}, Cell {1}", rowNum, columnNum)));

                                // write the end cell element
                                writer.WriteEndElement();
                            }

                            // write the end row element
                            writer.WriteEndElement();
                        }

                        // write the end SheetData element
                        writer.WriteEndElement();
                        // write the end Worksheet element
                        writer.WriteEndElement();
                        writer.Close();

                        writerXb.WriteElement(new Sheet
                        {
                            Name = dt.TableName,
                            SheetId = Convert.ToUInt32(count + 1),
                            Id = document.WorkbookPart.GetIdOfPart(workSheetPart)
                        });

                        count++;
                    }
                    // End Sheets
                    writerXb.WriteEndElement();
                    // End Workbook
                    writerXb.WriteEndElement();

                    writerXb.Close();

                    document.Close();
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                throw;
            }
        }

        [SuppressMessage("ReSharper", "PossiblyMistakenUseOfParamsMethod")]
        private static Stylesheet AddStyleSheet()
        {
            try
            {
                var workbookstylesheet = new Stylesheet();

                var font0 = new Font(); // Default font

                var font1 = new Font(); // Bold font
                var bold = new Bold();
                font1.Append(bold);

                var fonts = new Fonts(); // <APENDING Fonts>
                fonts.Append(font0);
                fonts.Append(font1);

                // <Fills>
                var fill0 = new Fill(); // Default fill

                var fills = new Fills(); // <APENDING Fills>
                fills.Append(fill0);

                // <Borders>
                var border0 = new Border(); // Defualt border

                var borders = new Borders(); // <APENDING Borders>
                borders.Append(border0);

                var nf2DateTime = new NumberingFormat
                {
                    NumberFormatId = UInt32Value.FromUInt32(7170),
                    FormatCode = StringValue.FromString("dd-MMM")
                };
                workbookstylesheet.NumberingFormats = new NumberingFormats();
                workbookstylesheet.NumberingFormats.Append(nf2DateTime);

                // <CellFormats>
                var cellformat0 = new CellFormat
                {
                    FontId = 0,
                    FillId = 0,
                    BorderId = 0
                }; // Default style : Mandatory | Style ID =0

                var cellformat1 = new CellFormat
                {
                    BorderId = 0,
                    FillId = 0,
                    FontId = 0,
                    NumberFormatId = 7170,
                    FormatId = 0,
                    ApplyNumberFormat = true
                };

                var cellformat2 = new CellFormat
                {
                    BorderId = 0,
                    FillId = 0,
                    FontId = 0,
                    NumberFormatId = 14,
                    FormatId = 0,
                    ApplyNumberFormat = true
                };

                // <APENDING CellFormats>
                var cellformats = new CellFormats();
                cellformats.Append(cellformat0);
                cellformats.Append(cellformat1);
                cellformats.Append(cellformat2);


                // Append FONTS, FILLS , BORDERS & CellFormats to stylesheet <Preserve the ORDER>
                workbookstylesheet.Append(fonts);
                workbookstylesheet.Append(fills);
                workbookstylesheet.Append(borders);
                workbookstylesheet.Append(cellformats);

                //// Finalize
                //stylesheet.Stylesheet = workbookstylesheet;
                //stylesheet.Stylesheet.Save();

                return workbookstylesheet;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                throw;
            }
        }
    }
}