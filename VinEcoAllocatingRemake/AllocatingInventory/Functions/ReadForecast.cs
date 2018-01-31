#region

#endregion

namespace VinEcoAllocatingRemake.AllocatingInventory
{
    #region

    using System;
    using System.Collections.Concurrent;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Data;
    using System.Diagnostics;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Threading.Tasks;

    using Aspose.Cells;

    using VinEcoAllocatingRemake.AllocatingInventory.Models;

    #endregion

    #region

    #endregion

    /// <summary>
    ///     The allocating inventory.
    /// </summary>
    public partial class AllocatingInventory
    {
        /// <summary>
        ///     Open External Config file ( Excel file )
        ///     to read and Update config.
        /// </summary>
        /// <param name="sender"> The sender. </param>
        /// <param name="e"> The e. </param>
        private async void ReadForecast(object sender, DoWorkEventArgs e)
        {
            try
            {
                var watch = new Stopwatch();
                watch.Start();

                var dicProduct = new ConcurrentDictionary<string, Product>();
                var dicSupplier = new ConcurrentDictionary<string, Supplier>();
                var dicFc = new ConcurrentDictionary<(DateTime DateFc, string ProductCode, string SupplierCode), (SupplierForecast Supply, bool Valid)>();
                var dicOldFc = new Dictionary<(DateTime DateFc, string ProductCode, string SupplierCode), (SupplierForecast Supply, bool Valid)>();

                this.WriteToRichTextBoxOutput(
                    "Đọc DBSL cũ từ cơ sở dữ liệu.",
                    1);

                // ReSharper disable ImplicitlyCapturedClosure
                // ReSharper disable HeapView.DelegateAllocation
                var readTasks = new[]
                                    {
                                        // Products
                                        new Task(
                                            delegate
                                                {
                                                    if (!File.Exists($@"{this._applicationPath}\Database\Products.xlsb")) return;

                                                    using (var workbook = new Workbook(
                                                        $@"{this._applicationPath}\Database\Products.xlsb",
                                                        new LoadOptions { MemorySetting = MemorySetting.MemoryPreference }))
                                                    {
                                                        Worksheet worksheet = workbook.Worksheets[0];
                                                        using (DataTable table = worksheet.Cells.ExportDataTable(
                                                            0,
                                                            0,
                                                            worksheet.Cells.MaxDataRow + 1,
                                                            worksheet.Cells.MaxDataColumn + 1,
                                                            this._globalExportTableOptionsOpts))
                                                        {
                                                            foreach (DataRow row in table.Select())
                                                                dicProduct.TryAdd(
                                                                    this._ulti.ObjectToString(row["ProductCode"]),
                                                                    new Product
                                                                        {
                                                                            ProductCode = this._ulti.ObjectToString(row["ProductCode"]),
                                                                            ProductName = this._ulti.ObjectToString(row["ProductName"])
                                                                        });
                                                        }
                                                    }
                                                }),

                                        // Suppliers
                                        new Task(
                                            delegate
                                                {
                                                    if (!File.Exists($@"{this._applicationPath}\Database\Suppliers.xlsb")) return;

                                                    using (var workbook = new Workbook(
                                                        $@"{this._applicationPath}\Database\Suppliers.xlsb",
                                                        new LoadOptions
                                                            {
                                                                MemorySetting =
                                                                    MemorySetting.MemoryPreference
                                                            }))
                                                    {
                                                        Worksheet worksheet = workbook.Worksheets[0];
                                                        using (DataTable table = worksheet.Cells.ExportDataTable(
                                                            0,
                                                            0,
                                                            worksheet.Cells.MaxDataRow + 1,
                                                            worksheet.Cells.MaxDataColumn + 1,
                                                            this._globalExportTableOptionsOpts))
                                                        {
                                                            foreach (DataRow row in table.Select())
                                                                dicSupplier.TryAdd(
                                                                    this._ulti.ObjectToString(row["SupplierCode"]),
                                                                    new Supplier
                                                                        {
                                                                            SupplierRegion = this._ulti.ObjectToString(row["SupplierRegion"]),
                                                                            SupplierType = this._ulti.ObjectToString(row["SupplierType"]),
                                                                            SupplierCode = this._ulti.ObjectToString(row["SupplierCode"]),
                                                                            SupplierName = this._ulti.ObjectToString(row["SupplierName"])
                                                                        });
                                                        }
                                                    }
                                                }),

                                        // Forecasts
                                        new Task(
                                            delegate
                                                {
                                                    if (!File.Exists($@"{this._applicationPath}\Database\Forecasts.xlsb")) return;

                                                    using (var workbook = new Workbook(
                                                        $@"{this._applicationPath}\Database\Forecasts.xlsb",
                                                        new LoadOptions
                                                            {
                                                                MemorySetting =
                                                                    MemorySetting.MemoryPreference
                                                            }))
                                                    {
                                                        Worksheet worksheet = workbook.Worksheets[0];
                                                        using (DataTable table = worksheet.Cells.ExportDataTable(
                                                            0,
                                                            0,
                                                            worksheet.Cells.MaxDataRow + 1,
                                                            worksheet.Cells.MaxDataColumn + 1,
                                                            this._globalExportTableOptionsOpts))
                                                        {
                                                            foreach (DataRow row in table.Select())
                                                            {
                                                                string productCode = this._ulti.ObjectToString(row["ProductCode"]);
                                                                string supplierCode = this._ulti.ObjectToString(row["SupplierCode"]);

                                                                for (var colIndex = 0;
                                                                     colIndex < table.Columns.Count;
                                                                     colIndex++)
                                                                    using (DataColumn column = table.Columns[colIndex])
                                                                    {
                                                                        // First check point. Is it a valid date?
                                                                        DateTime? dateFc = this._ulti.StringToDate(
                                                                            column.ColumnName);
                                                                        if (dateFc == null) continue;

                                                                        // Second check point. Is it a valid forecast value?
                                                                        double value = this._ulti.ObjectToDouble(row[colIndex]);
                                                                        if (value <= 0) continue;

                                                                        dicOldFc.Add(
                                                                            ((DateTime)dateFc, productCode, supplierCode),
                                                                            (new SupplierForecast { QualityControlPass = true, SupplierCode = supplierCode, FullOrder = this._ulti.ObjectToInt(row["FullOrder"]) == 1, CrossRegion = this._ulti.ObjectToInt(row["CrossRegion"]) == 1, LabelVinEco = this._ulti.ObjectToInt(row["Label"]) == 1, Level = (byte)this._ulti.ObjectToInt(row["Level"]) }, false));
                                                                    }
                                                            }
                                                        }
                                                    }
                                                })
                                    };

                // Here we go.
                Parallel.ForEach(
                    readTasks,
                    new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount },
                    task => { task.Start(); });

                await Task.WhenAll(readTasks).ConfigureAwait(true);

                // ReSharper restore HeapView.DelegateAllocation
                // ReSharper restore ImplicitlyCapturedClosure
                var listDt = new List<DataTable>();

                this.WriteToRichTextBoxOutput(
                    "Bắt đầu đọc DBSL mới.",
                    1);

                Parallel.ForEach(
                    new DirectoryInfo($@"{this._applicationPath}\Data\Forecast").GetFiles(),
                    new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount },
                    fileInfo =>
                        {
                            var stopwatch = new Stopwatch();
                            stopwatch.Start();

                            using (var workbook = new Workbook(
                                fileInfo.FullName,
                                new LoadOptions { MemorySetting = MemorySetting.MemoryPreference }))
                            {
                                Worksheet worksheet = workbook.Worksheets[0];

                                var rowIndex = 0;
                                var colIndex = 0;

                                // Initialize First value coz of While-loop.
                                string value = worksheet.Cells[rowIndex,
                                        colIndex]
                                    .Value?.ToString()
                                    .Trim();

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

                                    value = worksheet.Cells[rowIndex,
                                            colIndex]
                                        .Value?.ToString()
                                        .Trim();
                                }

                                // Core.
                                // Principle: Read all at once first.
                                // Then work on data later.
                                using (DataTable table = worksheet.Cells.ExportDataTable(
                                    rowIndex,
                                    colIndex,
                                    worksheet.Cells.MaxDataRow + 1,
                                    worksheet.Cells.MaxDataColumn + 1,
                                    this._globalExportTableOptionsOpts))
                                {
                                    table.TableName = Path.GetFileNameWithoutExtension(fileInfo.Name);

                                    // To deal with the uhm, Templates having different Headers.
                                    // Please shoot me.
                                    // ReSharper disable once SuggestVarOrType_SimpleTypes
                                    foreach (var key in new (string oldName, string newName)[]
                                                            {
                                                                ("Vùng", "Region"),
                                                                ("Mã Farm", "SCODE"),
                                                                ("Tên Farm", "SNAME"),
                                                                ("Nhóm", "PCLASS"),
                                                                ("Mã VECrops", "VECrops Code"),
                                                                ("Mã VinEco", "PCODE"),
                                                                ("Tên VinEco", "PNAME")
                                                            })
                                        if (table.Columns.Contains(key.oldName))
                                            table.Columns[key.oldName].ColumnName = key.newName;

                                    listDt.Add(table);
                                }
                            }

                            stopwatch.Stop();
                            this.WriteToRichTextBoxOutput(
                                $"{fileInfo.Name} - Xong trong {Math.Round(stopwatch.Elapsed.TotalSeconds, 2).ToString(CultureInfo.InvariantCulture)}s!",
                                2); // + " - Done!");
                        });

                this.WriteToRichTextBoxOutput(
                    "Bắt đầu xử lý DBSL.",
                    1);

                // Here comes the data handling.
                Parallel.ForEach(
                    listDt,
                    new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount },
                    table =>
                        {
                            try
                            {
                                bool isKpi = table.TableName.IndexOf(
                                                 "KPI",
                                                 StringComparison.OrdinalIgnoreCase)
                                             >= 0;

                                // Row layer.
                                foreach (DataRow row in table.Select())
                                {
                                    // Idk why this is a thing.
                                    // Empty or null ProductCode, yeah, sure, why not.
                                    if (string.IsNullOrEmpty(this._ulti.ObjectToString(row["PCODE"])))
                                    {
                                        // Just kidding, nope.
                                        continue;
                                    }

                                    // A simple function to return bool from Yes/No column.
                                    bool CheckCol(string colName, string comparer)
                                    {
                                        return table.Columns.Contains(colName)
                                               && string.Equals(
                                                   this._ulti.ObjectToString(row[colName]),
                                                   comparer,
                                                   StringComparison.OrdinalIgnoreCase);
                                    }

                                    // First check point.
                                    // If you are not allowed to go, well, see ya later.
                                    if (table.TableName != "VinEco" && !CheckCol("QC", "Ok") && !CheckCol("Source", "VinEco"))
                                    {
                                        continue;
                                    }

                                    // Less conversion.
                                    string supCode = this._ulti.ObjectToString(row["SCODE"]);
                                    string pCode = this._ulti.ObjectToString(row["PCODE"]);
                                    string pName = this._ulti.ObjectToString(row["PNAME"]).Replace("KH-", string.Empty);

                                    //// Product information.
                                    // Product product = dicProduct.GetOrAdd(pCode, new Product
                                    // {
                                    // ProductCode = pCode,
                                    // ProductName = _ulti.ObjectToString(row["PNAME"])
                                    // });
                                    dicProduct.AddOrUpdate(
                                        pCode,
                                        new Product
                                            {
                                                ProductCode = pCode,
                                                ProductName = pName
                                            },
                                        (key, oldProduct) =>
                                            new Product
                                                {
                                                    ProductCode = pCode,
                                                    ProductName =
                                                        string.CompareOrdinal(
                                                            oldProduct.ProductName,
                                                            pName)
                                                        < 0
                                                            ? pName
                                                            : oldProduct.ProductName
                                                });

                                    //// Quality of life. Get the pseudo 'best' Product Name.
                                    // if (string.CompareOrdinal(product.ProductName, _ulti.ObjectToString(row["PNAME"])) < 0)
                                    // product.ProductName = _ulti.ObjectToString(row["PNAME"]);

                                    // Optimization, dealing with region.
                                    string region = this._ulti.ObjectToString(row["Region"]);
                                    if (region.Contains(' ')) region = this._ulti.ConvertToUnsigned(this._ulti.ReturnInitials(region));

                                    // Supplier information.
                                    dicSupplier.TryAdd(
                                        supCode,
                                        new Supplier
                                            {
                                                SupplierRegion = region,
                                                SupplierCode = supCode,
                                                SupplierName = this._ulti.ObjectToString(row["SNAME"]),
                                                SupplierType = table.Columns.Contains("Source")
                                                                   ? this._ulti.ObjectToString(row["Source"])
                                                                   : table.TableName == "VinEco"
                                                                       ? "VinEco"
                                                                       : "ThuMua"
                                            });

                                    // Column layer.
                                    for (var colIndex = 0; colIndex < table.Columns.Count; colIndex++)
                                    {
                                        using (DataColumn column = table.Columns[colIndex])
                                        {
                                            // First check point. Is it a valid date?
                                            DateTime? dateFc = this._ulti.StringToDate(column.ColumnName);
                                            if (dateFc == null)
                                            {
                                                continue;
                                            }

                                            // Second check point. Is it a valid forecast value?
                                            // So this is causing a lot of freaking exception in ObjectToDouble.
                                            // Simply because it has to check too many things.
                                            double value = this._ulti.ObjectToDouble(row[colIndex]);
                                            if (value <= 0) continue;

                                            var supply = new SupplierForecast
                                                             {
                                                                 Availability =
                                                                     table.Columns.Contains("Availability")
                                                                         ? this._ulti.ObjectToString(row["Availability"])
                                                                         : "1234567",
                                                                 FullOrder = CheckCol("100%", "Yes"),
                                                                 LabelVinEco = CheckCol("Label VE", "Yes"),
                                                                 CrossRegion = CheckCol("CrossRegion", "Yes"),
                                                                 QualityControlPass = true,
                                                                 Level = (byte)(table.Columns.Contains("Level")
                                                                                    ? row["Level"] as byte? ?? 1
                                                                                    : 1),
                                                                 QuantityForecast = isKpi
                                                                                        ? 0
                                                                                        : value,
                                                                 QuantityForecastPlanned = !isKpi
                                                                                               ? 0
                                                                                               : value,
                                                                 HasKpi = isKpi
                                                             };

                                            // dicFc.GetOrAdd(((DateTime) dateFc, pCode, supCode),
                                            // (new SupplierForecast
                                            // {
                                            // SupplierCode = supCode,
                                            // Target = "All",
                                            // Availability = "1234567"
                                            // }, false)).Supply;
                                            dicFc.AddOrUpdate(
                                                ((DateTime)dateFc, pCode, supCode),
                                                (supply, false),
                                                (key, oldValue) =>
                                                    (new SupplierForecast
                                                         {
                                                             Availability = table.Columns.Contains("Availability")
                                                                                ? this._ulti.ObjectToString(row["Availability"])
                                                                                : "1234567",
                                                             FullOrder = CheckCol("100%", "Yes"),
                                                             LabelVinEco = CheckCol("Label VE", "Yes"),
                                                             CrossRegion = CheckCol("CrossRegion", "Yes"),
                                                             QualityControlPass = true,
                                                             Level = (byte)(table.Columns.Contains("Level")
                                                                                ? row["Level"] as byte? ?? 1
                                                                                : 1),
                                                             QuantityForecast = oldValue.Supply.QuantityForecast
                                                                                + (isKpi
                                                                                       ? 0
                                                                                       : value),
                                                             QuantityForecastPlanned = oldValue.Supply.QuantityForecast
                                                                                       + (!isKpi
                                                                                              ? 0
                                                                                              : value),
                                                             HasKpi = isKpi
                                                         }, false));

                                            // var myLock = new object();
                                            // lock (myLock)
                                            // {
                                            // if (isKpi)
                                            // {
                                            // supply.HasKpi = true;
                                            // supply.QuantityForecastPlanned += value;
                                            // }
                                            // else
                                            // {
                                            // supply.QuantityForecast += value;
                                            // }
                                            // }
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                this.WriteToRichTextBoxOutput(ex.Message);
                                throw;
                            }
                        });

                this.WriteToRichTextBoxOutput(
                    $"Xử lý xong DBSL, mất: {Math.Round(watch.Elapsed.TotalSeconds, 2).ToString(CultureInfo.InvariantCulture)}s!",
                    2);

                // ReSharper disable ImplicitlyCapturedClosure
                // ReSharper disable HeapView.DelegateAllocation
                var writeTasks = new[]
                                     {
                                         // Forecasts
                                         new Task(
                                             delegate
                                                 {
                                                     try
                                                     {
                                                         using (var table = new DataTable { TableName = "Forecasts" })
                                                         {
                                                             // ReSharper disable once SuggestVarOrType_SimpleTypes
                                                             foreach (var key in new (string colName, Type colType)[]
                                                                                     {
                                                                                         ("ProductCode", typeof(string)),
                                                                                         ("SupplierCode", typeof(string)),
                                                                                         ("FullOrder", typeof(int)),
                                                                                         ("Label", typeof(int)),
                                                                                         ("CrossRegion", typeof(int)),
                                                                                         ("Level", typeof(int))
                                                                                     })
                                                                 table.Columns.Add(key.colName, key.colType);

                                                             var listDateFc = new List<DateTime>();
                                                             var listAllDateFc = new List<DateTime>();

                                                             // Count DateFc.
                                                             // ReSharper disable once SuggestVarOrType_SimpleTypes
                                                             foreach (var key in dicFc.Keys)
                                                             {
                                                                 if (!listDateFc.Contains(key.DateFc)) listDateFc.Add(key.DateFc);

                                                                 if (!listAllDateFc.Contains(key.DateFc)) listAllDateFc.Add(key.DateFc);
                                                             }

                                                             // ... and then add the same amount of columns.
                                                             foreach (DateTime dateFc in listDateFc)
                                                                 // ReSharper disable once SuggestVarOrType_SimpleTypes
                                                             foreach (var key in dicOldFc.Keys.ToList())
                                                             {
                                                                 if (key.DateFc == dateFc) dicOldFc.Remove(key);

                                                                 // Also remove all old items.
                                                                 if (!listAllDateFc.Contains(key.DateFc)) listAllDateFc.Add(key.DateFc);
                                                             }

                                                             foreach (DateTime dateFc in listAllDateFc.OrderBy(d => d.Date)) table.Columns.Add(this._ulti.DateToString(dateFc, "dd-MMM-yyyy"), typeof(double));

                                                             // Dictionary of rowIndex.
                                                             var dicRow = new Dictionary<string, int>(
                                                                 dicProduct.Count,
                                                                 StringComparer.OrdinalIgnoreCase);

                                                             object objIntOne = this._ulti.IntToObject(1);

                                                             object BoolToOne(bool expression)
                                                             {
                                                                 return expression
                                                                            ? objIntOne
                                                                            : DBNull.Value;
                                                             }

                                                             // Hour of truth.
                                                             // ReSharper disable once SuggestVarOrType_SimpleTypes
                                                             foreach (var key in dicFc.Keys) dicOldFc.Add(key, dicFc[key]);

                                                             foreach (
                                                                 // ReSharper disable once SuggestVarOrType_SimpleTypes
                                                                 var key in from key in dicOldFc.Keys
                                                                            orderby key.ProductCode, key
                                                                                .SupplierCode
                                                                            select key)
                                                             {
                                                                 DataRow row;

                                                                 string rowKey = $"{key.ProductCode}{key.SupplierCode}";
                                                                 SupplierForecast supply =
                                                                     dicOldFc[(key.DateFc, key.ProductCode, key.SupplierCode)]
                                                                         .Supply;

                                                                 if (dicRow.TryGetValue(rowKey, out int rowIndex))
                                                                 {
                                                                     row = table.Select()[rowIndex];
                                                                 }
                                                                 else
                                                                 {
                                                                     row = table.NewRow();

                                                                     row["ProductCode"] = key.ProductCode;
                                                                     row["SupplierCode"] = key.SupplierCode;
                                                                     row["FullOrder"] = BoolToOne(supply.FullOrder);
                                                                     row["Label"] = BoolToOne(supply.LabelVinEco);
                                                                     row["CrossRegion"] = BoolToOne(supply.CrossRegion);
                                                                     row["Level"] = this._ulti.IntToObject(supply.Level);

                                                                     dicRow.Add(rowKey, table.Rows.Count);
                                                                     table.Rows.Add(row);
                                                                 }

                                                                 row[this._ulti.DateToString(
                                                                     key.DateFc,
                                                                     "dd-MMM-yyyy")] = this._ulti.DoubleToObject(
                                                                     supply.QuantityForecast);
                                                             }

                                                             string path =
                                                                 $@"{this._applicationPath}\Database\{table.TableName}.xlsx";
                                                             this._ulti.LargeExportOneWorkbook(
                                                                 path,
                                                                 new List<DataTable> { table },
                                                                 true,
                                                                 true);
                                                             this._ulti.ConvertExcelTypeInterop(
                                                                 path,
                                                                 "xlsx",
                                                                 "xlsb");
                                                         }
                                                     }
                                                     catch (Exception ex)
                                                     {
                                                         this.WriteToRichTextBoxOutput(ex.Message);
                                                         throw;
                                                     }
                                                 }),

                                         // Products
                                         new Task(
                                             delegate
                                                 {
                                                     using (var table = new DataTable { TableName = "Products" })
                                                     {
                                                         // ReSharper disable once SuggestVarOrType_SimpleTypes
                                                         foreach (var key in new (string colName, Type colType)[]
                                                                                 {
                                                                                     ("ProductCode", typeof(string)),
                                                                                     ("ProductName", typeof(string))
                                                                                 })
                                                             table.Columns.Add(key.colName, key.colType);

                                                         foreach (Product product in from value in dicProduct.Values
                                                                                     orderby value.ProductCode
                                                                                     select value)
                                                         {
                                                             DataRow row = table.NewRow();

                                                             row["ProductCode"] = product.ProductCode;
                                                             row["ProductName"] = product.ProductName;

                                                             table.Rows.Add(row);
                                                         }

                                                         string path = $@"{this._applicationPath}\Database\{table.TableName}.xlsx";
                                                         this._ulti.LargeExportOneWorkbook(
                                                             path,
                                                             new List<DataTable> { table },
                                                             true,
                                                             true);
                                                         this._ulti.ConvertExcelTypeInterop(
                                                             path,
                                                             "xlsx",
                                                             "xlsb");
                                                     }
                                                 }),

                                         // Suppliers
                                         new Task(
                                             delegate
                                                 {
                                                     using (var table = new DataTable { TableName = "Suppliers" })
                                                     {
                                                         // ReSharper disable once SuggestVarOrType_SimpleTypes
                                                         foreach (var key in new (string colName, Type colType)[]
                                                                                 {
                                                                                     ("SupplierType", typeof(string)),
                                                                                     ("SupplierRegion", typeof(string)),
                                                                                     ("SupplierCode", typeof(string)),
                                                                                     ("SupplierName", typeof(string))
                                                                                 })
                                                             table.Columns.Add(key.colName, key.colType);

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

                                                         string path = $@"{this._applicationPath}\Database\{table.TableName}.xlsx";
                                                         this._ulti.LargeExportOneWorkbook(
                                                             path,
                                                             new List<DataTable> { table },
                                                             true,
                                                             true);
                                                         this._ulti.ConvertExcelTypeInterop(
                                                             path,
                                                             "xlsx",
                                                             "xlsb");
                                                     }
                                                 })
                                     };

                // Here we go.
                Parallel.ForEach(
                    writeTasks,
                    new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount },
                    task => { task.Start(); });

                // Making sure every Tasks finished before proceeding.
                await Task.WhenAll(writeTasks).ConfigureAwait(true);

                // The final flag.
                watch.Stop();
                this.WriteToRichTextBoxOutput(
                    $"Đã ghi vào cơ sở dữ liệu. Tổng thời gian chạy: {Math.Round(watch.Elapsed.TotalSeconds, 2).ToString(CultureInfo.InvariantCulture)}s!",
                    1);
            }
            catch (Exception ex)
            {
                // Just, why?
                this.WriteToRichTextBoxOutput(ex.Message);
                throw;
            }
            finally
            {
                this.TryClear();
            }
        }
    }
}