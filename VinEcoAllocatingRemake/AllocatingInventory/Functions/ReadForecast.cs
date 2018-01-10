using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Aspose.Cells;
using VinEcoAllocatingRemake.AllocatingInventory.Models;

namespace VinEcoAllocatingRemake.AllocatingInventory
{
    public partial class AllocatingInventory
    {
        /// <summary>
        ///     Open External Config file ( Excel file )
        ///     to read and Update config.
        /// </summary>
        [SuppressMessage("ReSharper", "ArgumentsStyleLiteral")]
        [SuppressMessage("ReSharper", "ArgumentsStyleOther")]
        [SuppressMessage("ReSharper", "ArgumentsStyleNamedExpression")]
        [SuppressMessage("ReSharper", "ArgumentsStyleStringLiteral")]
        private void ReadForecast(object sender, DoWorkEventArgs e)
        {
            try
            {
                var watch = new Stopwatch();
                watch.Start();

                #region Initializing variables

                var dicProduct = new ConcurrentDictionary<string, Product>();
                var dicSupplier = new ConcurrentDictionary<string, Supplier>();
                var dicFc =
                    new ConcurrentDictionary<(DateTime DateFc, string ProductCode, string SupplierCode), (
                        SupplierForecast Supply, bool)>();

                #endregion

                WriteToRichTextBoxOutput("Đọc DBSL cũ từ cơ sở dữ liệu.", 1);

                #region Reading old data.

                // Todo - oh my fucking god.

                #endregion

                var listDt = new List<DataTable>();

                WriteToRichTextBoxOutput("Bắt đầu đọc DBSL mới.", 1);

                #region Reading new data.

                Parallel.ForEach(new DirectoryInfo($@"{_applicationPath}\Data\Forecast").GetFiles(),
                    new ParallelOptions {MaxDegreeOfParallelism = Environment.ProcessorCount},
                    fileInfo =>
                    {
                        var stopwatch = new Stopwatch();
                        stopwatch.Start();

                        var opt = new LoadOptions {MemorySetting = MemorySetting.MemoryPreference};
                        var xlWb = new Workbook(fileInfo.FullName, opt);
                        //Worksheet xlWs = xlWb.Worksheets.OrderByDescending(x => x.Cells.MaxDataRow).First();

                        // Convert LinQ to code.
                        // Avoid sorting.
                        Worksheet xlWs = xlWb.Worksheets[0];
                        //foreach (Worksheet ws in xlWb.Worksheets.ToList())
                        //    if (ws.Cells.MaxDataRow > xlWs.Cells.MaxDataRow)
                        //        xlWs = ws;

                        var rowIndex = 0;
                        var colIndex = 0;

                        // Initialize First value coz of While-loop.
                        string value = xlWs.Cells[rowIndex, colIndex].Value?.ToString().Trim();

                        // Search for the very first row.
                        while (value != "Vùng" && value != "Region" && rowIndex <= 100 && colIndex <= 100)
                        {
                            // Next row.
                            rowIndex++;

                            // If above 100.
                            if (rowIndex > 100)
                            {
                                rowIndex = 0;
                                colIndex++;
                            }

                            // Checkpoint. Well, there has to be a limit.
                            if (colIndex > 100) break;

                            value = xlWs.Cells[rowIndex, colIndex].Value?.ToString().Trim();
                        }

                        // Core.
                        // Principle: Read all at once first.
                        // Then work on data later.
                        using (DataTable table = xlWs.Cells.ExportDataTable(rowIndex, colIndex,
                            xlWs.Cells.MaxDataRow + 1,
                            xlWs.Cells.MaxDataColumn + 1, _globalExportTableOptionsopts))
                        {
                            table.TableName = Path.GetFileNameWithoutExtension(fileInfo.Name);
                            // To deal with the uhm, Templates having different Headers.
                            // Please shoot me.

                            foreach ((string oldName, string newName) in new[]
                            {
                                ("Vùng", "Region"),
                                ("Mã Farm", "SCODE"),
                                ("Tên Farm", "SNAME"),
                                ("Nhóm", "PCLASS"),
                                ("Mã VECrops", "VECrops Code"),
                                ("Mã VinEco", "PCODE"),
                                ("Tên VinEco", "PNAME")
                            })
                                if (table.Columns.Contains(oldName))
                                    table.Columns[oldName].ColumnName = newName;

                            listDt.Add(table);
                        }

                        stopwatch.Stop();
                        WriteToRichTextBoxOutput(
                            message:
                            $"{fileInfo.Name} - Xong trong {Math.Round(stopwatch.Elapsed.TotalSeconds, 2).ToString(CultureInfo.InvariantCulture)}s!",
                            importanceLevel: 2); // + " - Done!");
                    });

                #endregion

                WriteToRichTextBoxOutput("Bắt đầu xử lý DBSL.", 1);

                #region Handling Data.

                // Here comes the data handling.
                Parallel.ForEach(listDt, new ParallelOptions {MaxDegreeOfParallelism = Environment.ProcessorCount},
                    table =>
                    {
                        try
                        {
                            bool isKpi = table.TableName.IndexOf("KPI", StringComparison.OrdinalIgnoreCase) >= 0;
                            // Row layer.
                            foreach (DataRow row in table.Select())
                            {
                                // Idk why this is a thing.
                                if (string.IsNullOrEmpty(_ulti.ObjectToString(row["PCODE"]))) continue;

                                bool CheckCol(string colName, string comparer)
                                {
                                    return table.Columns.Contains(colName) &&
                                           string.Equals(
                                               _ulti.ObjectToString(row[colName]),
                                               comparer, StringComparison.OrdinalIgnoreCase);
                                }

                                // First check point.
                                // If you are not allowed to go, well, see ya later.
                                if (table.TableName != "VinEco" &&
                                    !CheckCol("QC", "Ok") &&
                                    !CheckCol("Source", "VinEco")) continue;

                                // Less conversion.
                                string supCode = _ulti.ObjectToString(row["SCODE"]);
                                string pCode = _ulti.ObjectToString(row["PCODE"]);

                                //Product product = dicProduct.GetOrAdd(pCode, new Product
                                //{
                                //    ProductCode = pCode,
                                //    ProductName = _ulti.ObjectToString(row["PNAME"]),
                                //});

                                dicProduct.TryAdd(pCode, new Product
                                {
                                    ProductCode = pCode,
                                    ProductName = _ulti.ObjectToString(row["PNAME"])
                                });

                                //Supplier supplier = dicSupplier.GetOrAdd(supCode, new Supplier
                                //{
                                //    SupplierRegion = _ulti.ObjectToString(row["Region"]),
                                //    SupplierCode = supCode,
                                //    SupplierName = _ulti.ObjectToString(row["SNAME"]),
                                //});

                                dicSupplier.TryAdd(supCode, new Supplier
                                {
                                    SupplierRegion = _ulti.ObjectToString(row["Region"]),
                                    SupplierCode = supCode,
                                    SupplierName = _ulti.ObjectToString(row["SNAME"]),
                                    SupplierType = table.Columns.Contains("Source")
                                        ? _ulti.ObjectToString(row["Source"])
                                        : table.TableName == "VinEco"
                                            ? "VinEco"
                                            : "ThuMua"
                                });

                                // Column layer.
                                for (var colIndex = 0; colIndex < table.Columns.Count; colIndex++)
                                    using (DataColumn colummn = table.Columns[colIndex])
                                    {
                                        // First check point. Is it a valid date?
                                        DateTime? dateFc = _ulti.StringToDate(colummn.ColumnName);
                                        if (dateFc == null) continue;

                                        // Second check point. Is it a valid forecast value?
                                        double fcValue = _ulti.ObjectToDouble(row[colIndex]);
                                        if (fcValue <= 0) continue;

                                        SupplierForecast supply = dicFc.GetOrAdd(((DateTime) dateFc, pCode, supCode),
                                            (new SupplierForecast
                                            {
                                                SupplierCode = supCode
                                            }, false)).Supply;

                                        supply.Availability = table.Columns.Contains("Availability")
                                            ? _ulti.ObjectToString("Availability")
                                            : "1234567";

                                        supply.FullOrder = CheckCol("100%", "Yes");
                                        supply.LabelVinEco = CheckCol("Label VE", "Yes");
                                        supply.CrossRegion = CheckCol("CrossRegion", "Yes");
                                        supply.QualityControlPass = true;
                                        supply.Level = (byte) (table.Columns.Contains("Level")
                                            ? row["Level"] as byte? ?? 1
                                            : 1);

                                        lock (supply)
                                        {
                                            if (isKpi)
                                            {
                                                supply.HasKpi = true;
                                                supply.QuantityForecastPlanned += fcValue;
                                            }
                                            else
                                            {
                                                supply.QuantityForecast += fcValue;
                                            }
                                        }
                                    }
                            }
                        }
                        catch (Exception ex)
                        {
                            WriteToRichTextBoxOutput(ex.Message);
                            throw;
                        }
                    });

                #endregion

                watch.Stop();
                WriteToRichTextBoxOutput(
                    $"Xử lý xong DBSL, Tổng thời gian: {Math.Round(watch.Elapsed.TotalSeconds, 2).ToString(CultureInfo.InvariantCulture)}s!",
                    2);

                // ReSharper disable ImplicitlyCapturedClosure
                // ReSharper disable HeapView.DelegateAllocation

                // Forecasts
                Task.Run(delegate
                {
                    try
                    {
                        using (var table = new DataTable {TableName = "Forecasts"})
                        {
                            foreach ((string colName, Type colType) in new[]
                            {
                                ("ProductCode", typeof(string)),
                                ("SupplierCode", typeof(string))
                            })
                            {
                                table.Columns.Add(colName, colType);
                            }

                            var listDateFc = new List<DateTime>();

                            foreach ((DateTime dateFc, string _, string _) in dicFc.Keys)
                            {
                                if (!listDateFc.Contains(dateFc))
                                    listDateFc.Add(dateFc);
                            }

                            foreach (DateTime dateFc in listDateFc.OrderBy(x => x))
                            {
                                table.Columns.Add(_ulti.DateToString(dateFc, "dd-MMM-yyyy"), typeof(double));
                            }

                            var dicRow =
                                new Dictionary<string, int>(dicProduct.Count, StringComparer.OrdinalIgnoreCase);

                            foreach ((DateTime dateFc, string productCode, string supplierCode) in 
                                from key in dicFc.Keys
                                orderby key.ProductCode, key.SupplierCode
                                select key)
                            {
                                DataRow row;

                                string rowKey = $"{productCode}{supplierCode}";

                                if (dicRow.TryGetValue(rowKey, out int rowIndex))
                                    row = table.Rows[rowIndex];
                                else
                                {
                                    row = table.NewRow();

                                    row["ProductCode"] = productCode;
                                    row["SupplierCode"] = supplierCode;

                                    dicRow.Add(rowKey, table.Rows.Count);
                                    table.Rows.Add(row);
                                }

                                SupplierForecast supply = dicFc[(dateFc, productCode, supplierCode)].Supply;

                                row[_ulti.DateToString(dateFc, "dd-MMM-yyyy")] =
                                    _ulti.DoubleToObject(supply.QuantityForecast);
                            }

                            LargeExportOneWorkbook($@"{_applicationPath}\Database\{table.TableName}.xlsx",
                                new List<DataTable> {table}, true, true);
                        }
                    }
                    catch (Exception ex)
                    {
                        WriteToRichTextBoxOutput(ex.Message);
                        throw;
                    }
                });

                // Products
                Task.Run(delegate
                {
                    using (var table = new DataTable {TableName = "Products"})
                    {
                        foreach ((string colName, Type colType) in new[]
                        {
                            ("ProductCode", typeof(string)),
                            ("ProductName", typeof(string))
                        })
                        {
                            table.Columns.Add(colName, colType);
                        }

                        foreach (Product product in
                            from value in dicProduct.Values
                            orderby value.ProductCode
                            select value)
                        {
                            DataRow row = table.NewRow();

                            row["ProductCode"] = product.ProductCode;
                            row["ProductName"] = product.ProductName;

                            table.Rows.Add(row);
                        }

                        LargeExportOneWorkbook($@"{_applicationPath}\Database\{table.TableName}.xlsx",
                            new List<DataTable> {table}, true, true);
                    }
                });

                // Suppliers
                Task.Run(delegate
                {
                    using (var table = new DataTable {TableName = "Suppliers"})
                    {
                        foreach ((string colName, Type colType) in new[]
                        {
                            ("SupplierType", typeof(string)),
                            ("SupplierRegion", typeof(string)),
                            ("SupplierCode", typeof(string)),
                            ("SupplierName", typeof(string)),
                        })
                        {
                            table.Columns.Add(colName, colType);
                        }

                        foreach (Supplier supplier in
                            from supplier in dicSupplier.Values
                            orderby
                                supplier.SupplierType,
                                supplier.SupplierRegion,
                                supplier.SupplierCode
                            select supplier)
                        {
                            DataRow row = table.NewRow();

                            row["SupplierType"] = supplier.SupplierType;
                            row["SupplierRegion"] = supplier.SupplierRegion;
                            row["SupplierCode"] = supplier.SupplierCode;
                            row["SupplierName"] = supplier.SupplierName;

                            table.Rows.Add(row);
                        }

                        LargeExportOneWorkbook($@"{_applicationPath}\Database\{table.TableName}.xlsx",
                            new List<DataTable> {table}, true, true);
                    }
                });

                // ReSharper restore HeapView.DelegateAllocation
                // ReSharper restore ImplicitlyCapturedClosure
            }
            // Just, why?
            catch (Exception ex)
            {
                WriteToRichTextBoxOutput(ex.Message);
                throw;
            }
        }
    }
}