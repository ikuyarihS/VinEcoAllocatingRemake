namespace VinEcoAllocatingRemake.AllocatingInventory
{
    #region

    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Globalization;

    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;

    #endregion

    #region

    #endregion

    /// <summary>
    ///     The utilities.
    /// </summary>
    public partial class Utilities
    {
        /// <summary>
        ///     OpenWriter Style, for Multiple DataTable into Multiple Worksheets in a single Workbook. A real fucking pain.
        /// </summary>
        /// <param name="filePath">Where your file will be.</param>
        /// <param name="theName">Oh come on.</param>
        /// <param name="listArrays">List of Arrays. Can contain just 1, doesn't matter.</param>
        /// <param name="listColumnNames">Ah fuck.</param>
        /// <param name="listTypes">Ah fuck v2.</param>
        /// <param name="yesHeader">You want headers?</param>
        /// <param name="yesZero">You want zero instead of null?</param>
        public void ExportXmlArray(
            string filePath,
            string theName,
            IEnumerable<object[,]> listArrays,
            List<string> listColumnNames,
            List<Type> listTypes,
            bool yesHeader = false,
            bool yesZero = false)
        {
            try
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Create(
                    filePath,
                    SpreadsheetDocumentType.Workbook))
                {
                    document.AddWorkbookPart();

                    OpenXmlWriter writerXb = OpenXmlWriter.Create(document.WorkbookPart);
                    writerXb.WriteStartElement(new Workbook());
                    writerXb.WriteStartElement(new Sheets());

                    var count = 0;

                    foreach (object[,] array in listArrays)
                    {
                        var dicColName = new Dictionary<int, string>();

                        for (var colIndex = 0; colIndex < array.GetUpperBound(1); colIndex++)
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
                                              { typeof(DateTime), "Date" },
                                              { typeof(string), "InlineString" },
                                              { typeof(double), "Number" },
                                              { typeof(int), "Number" },
                                              { typeof(bool), "Boolean" }
                                          };

                        // this list of attributes will be used when writing a start element
                        List<OpenXmlAttribute> attributes;

                        var workSheetPart = document.WorkbookPart.AddNewPart<WorksheetPart>();

                        OpenXmlWriter writer = OpenXmlWriter.Create(workSheetPart);
                        writer.WriteStartElement(new Worksheet());
                        writer.WriteStartElement(new SheetData());

                        if (yesHeader)
                        {
                            // create a new list of attributes
                            attributes = new List<OpenXmlAttribute>
                                             {
                                                 // add the row index attribute to the list
                                                 new OpenXmlAttribute("r", null, 1.ToString())
                                             };

                            // write the row start element with the row index attribute
                            writer.WriteStartElement(new Row(), attributes);

                            for (var columnNum = 1; columnNum <= array.GetUpperBound(1); ++columnNum)
                            {
                                // reset the list of attributes
                                attributes = new List<OpenXmlAttribute>
                                                 {
                                                     new OpenXmlAttribute("t", null, "str"),
                                                     new OpenXmlAttribute(
                                                         "r",
                                                         string.Empty,
                                                         $"{dicColName[columnNum]}1")
                                                 };

                                // add data type attribute - in this case inline string (you might want to look at the shared strings table)
                                // add the cell reference attribute

                                // write the cell start element with the type and reference attributes
                                writer.WriteStartElement(new Cell(), attributes);

                                // write the cell value
                                writer.WriteElement(new CellValue(listColumnNames[columnNum - 1]));

                                // writer.WriteElement(new CellValue(string.Format("This is Row {0}, Cell {1}", rowNum, columnNum)));

                                // write the end cell element
                                writer.WriteEndElement();
                            }

                            // write the end row element
                            writer.WriteEndElement();
                        }

                        for (var rowNum = 1; rowNum <= array.GetUpperBound(0); rowNum++)
                        {
                            // create a new list of attributes
                            attributes = new List<OpenXmlAttribute> { new OpenXmlAttribute("r", null, (yesHeader ? rowNum + 1 : rowNum).ToString()) };

                            // add the row index attribute to the list

                            // write the row start element with the row index attribute
                            writer.WriteStartElement(new Row(), attributes);

                            // DataRow dr = dt.Rows[rowNum - 1];
                            for (var columnNum = 1; columnNum <= array.GetUpperBound(1); columnNum++)
                            {
                                Type type = listTypes[columnNum - 1];

                                // reset the list of attributes
                                attributes = new List<OpenXmlAttribute>
                                                 {
                                                     // Add data type attribute - in this case inline string (you might want to look at the shared strings table)
                                                     new OpenXmlAttribute("t", null, type == typeof(string) ? "str" : dicType[type]),

                                                     // Add the cell reference attribute
                                                     new OpenXmlAttribute("r", string.Empty, $"{dicColName[columnNum]}{(yesHeader ? rowNum + 1 : rowNum).ToString(CultureInfo.InvariantCulture)}")
                                                 };

                                // write the cell start element with the type and reference attributes
                                writer.WriteStartElement(new Cell(), attributes);

                                // write the cell value
                                if (!yesZero && array[rowNum - 1, columnNum - 1] != null && array[rowNum - 1, columnNum - 1].ToString() != "0")
                                    writer.WriteElement(new CellValue(array[rowNum - 1, columnNum - 1].ToString()));

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

                        writerXb.WriteElement(
                            new Sheet
                                {
                                    Name = theName,
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