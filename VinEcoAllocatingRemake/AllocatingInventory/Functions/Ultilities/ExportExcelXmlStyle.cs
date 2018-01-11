using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace VinEcoAllocatingRemake.AllocatingInventory
{
    public partial class Utilities
    {
        /// <summary>
        ///     OpenWriter Style, for Multiple DataTable into Multiple Worksheets in a single Workbook. A real fucking pain.
        /// </summary>
        /// <param name="filePath">Where your file will be.</param>
        /// <param name="listDataTables">List of dataTables. Can contain just 1, doesn't matter.</param>
        /// <param name="yesHeader">You want headers?</param>
        /// <param name="yesZero">You want zero instead of null?</param>
        public void LargeExportOneWorkbook(
            string filePath,
            IEnumerable<DataTable> listDataTables,
            bool yesHeader = false,
            bool yesZero = false)
        {
            try
            {
                using (SpreadsheetDocument document =
                    SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
                {
                    document.AddWorkbookPart();

                    OpenXmlWriter writerXb = OpenXmlWriter.Create(document.WorkbookPart);
                    writerXb.WriteStartElement(new Workbook());
                    writerXb.WriteStartElement(new Sheets());

                    var count = 0;

                    foreach (DataTable dt in listDataTables)
                    {
                        var dicColName = new Dictionary<int, string>();

                        for (var colIndex = 0; colIndex < dt.Columns.Count; colIndex++)
                        {
                            int dividend = colIndex + 1;
                            string columnName = string.Empty;

                            while (dividend > 0)
                            {
                                int modifier = (dividend - 1) % 26;
                                columnName =
                                    $"{Convert.ToChar(65 + modifier).ToString(CultureInfo.InvariantCulture)}{columnName}";
                                dividend = (dividend - modifier) / 26;
                            }

                            dicColName.Add(colIndex + 1, columnName);
                        }

                        var dicType = new Dictionary<Type, string>(4)
                        {
                            // Neccessary evil.
                            //{typeof(DateTime), CellValues.Date.ToString()},
                            //{typeof(string), CellValues.InlineString.ToString()},
                            //{typeof(double), CellValues.Number.ToString()},
                            //{typeof(int), CellValues.Number.ToString()},
                            //{typeof(bool), CellValues.Boolean.ToString()}
                            {typeof(DateTime), "Date"},
                            {typeof(string), "InlineString"},
                            {typeof(double), "Number"},
                            {typeof(int), "Number"},
                            {typeof(bool), "Boolean"}
                        };

                        //this list of attributes will be used when writing a start element
                        List<OpenXmlAttribute> attributes;

                        var workSheetPart = document.WorkbookPart.AddNewPart<WorksheetPart>();

                        OpenXmlWriter writer = OpenXmlWriter.Create(workSheetPart);
                        writer.WriteStartElement(new Worksheet());
                        writer.WriteStartElement(new SheetData());

                        if (yesHeader)
                        {
                            //create a new list of attributes
                            attributes = new List<OpenXmlAttribute>
                            {
                                // add the row index attribute to the list
                                new OpenXmlAttribute("r", null, 1.ToString())
                            };

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
                                    (yesHeader ? rowNum + 1 : rowNum).ToString())
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
                                    // Add data type attribute - in this case inline string (you might want to look at the shared strings table)
                                    new OpenXmlAttribute("t", null,
                                        type == typeof(string) ? "str" : dicType[type]),
                                    // Add the cell reference attribute
                                    new OpenXmlAttribute("r", "",
                                        $"{dicColName[columnNum]}{(yesHeader ? rowNum + 1 : rowNum).ToString(CultureInfo.InvariantCulture)}")
                                };

                                //write the cell start element with the type and reference attributes
                                writer.WriteStartElement(new Cell(), attributes);

                                //write the cell value
                                if (yesZero | (dr[columnNum - 1].ToString() != "0"))
                                    writer.WriteElement(new CellValue(dr[columnNum - 1].ToString()));

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
    }
}