using System.Data;
using Aspose.Cells;

namespace VinEcoAllocatingRemake.AllocatingInventory
    {
        #region

        #endregion

        /// <summary>
        ///     The utilities.
        /// </summary>
        public partial class Utilities
            {
                /// <summary>
                ///     Aspose.Cells Approach. Now with MemoryPreference
                /// </summary>
                /// <param name="dataTable"> The data Table. </param>
                /// <param name="workbook"> The workbook. </param>
                /// <param name="yesNoHeader"> The yes No Header. </param>
                /// <param name="rowFirst"> The row First. </param>
                /// <param name="customDateFormat"> The custom Date Format. </param>
                public void OutputExcelAspose(
                    DataTable dataTable,
                    Workbook  workbook,
                    bool      yesNoHeader      = false,
                    int       rowFirst         = 6,
                    string    customDateFormat = "")
                    {
                        Style defaultStyle = workbook.CreateStyle();

                        defaultStyle.Font.Name = "Calibri";
                        defaultStyle.Font.Size = 11;

                        workbook.DefaultStyle = defaultStyle;

                        int rowTotal = dataTable.Rows.Count;
                        int colTotal = dataTable.Columns.Count;

                        // Add new worksheet.
                        Worksheet worksheet = workbook.Worksheets.Add(dataTable.TableName);

                        // Optimize for Performance?
                        worksheet.Cells.MemorySetting = MemorySetting.MemoryPreference;

                        // Import DataTable into worksheet.
                        worksheet.Cells.ImportDataTable(
                            dataTable,
                            yesNoHeader,
                            rowFirst - 1,
                            0,
                            rowTotal,
                            colTotal,
                            false,
                            customDateFormat == string.Empty ? "dd-MMM" : customDateFormat,
                            false);

                        // Set AutoFilter Range.
                        worksheet.AutoFilter.Range =
                            $"A1:{worksheet.Cells[worksheet.Cells.MaxDataRow + 1, worksheet.Cells.MaxDataColumn + 1].Name}";
                    }
            }
    }