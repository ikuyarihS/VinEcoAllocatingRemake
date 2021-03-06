﻿// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ReadPurchaseOrder.cs" company="VinEco">
//   Shirayuki 2018.
// </copyright>
// <summary>
//   The allocating inventory.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

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
        // ReSharper disable once StyleCop.SA1404
        [SuppressMessage("ReSharper", "ArrangeThisQualifier")]
        private async void ReadPurchaseOrder(object sender, DoWorkEventArgs e)
        {
            try
            {
                var watch = new Stopwatch();
                watch.Start();

                var dicProduct  = new ConcurrentDictionary<string, Product>();
                var dicCustomer = new ConcurrentDictionary<string, Customer>();
                var dicPo       = new ConcurrentDictionary<(DateTime DatePo, string ProductCode, string CustomerKeyCode), (CustomerOrder Order, bool)>();

                // var dicOldPo =
                // new Dictionary<(DateTime DateFc, string ProductCode, string CustomerKeyCode),
                // (CustomerOrder Order, bool)>();

                // Todo - Implement this instead of a Dictionary because resizing is being a bitch.
                var listOldPo = new List<((DateTime datePo, string ProductCode, string CustomerKeyCode) Key, (CustomerOrder Order, bool) Value)>();

                this.WriteToRichTextBoxOutput("Đọc Đơn hàng cũ từ cơ sở dữ liệu.", 1);

                var readTasks = new[]
                                    {
                                        // Products
                                        new Task(
                                            delegate
                                            {
                                                if (!File.Exists($@"{this.applicationPath}\Database\Products.xlsb")) return;

                                                using (var workbook = new Workbook(
                                                    $@"{this.applicationPath}\Database\Products.xlsb",
                                                    new LoadOptions { MemorySetting = MemorySetting.MemoryPreference }))
                                                {
                                                    Worksheet worksheet = workbook.Worksheets[0];
                                                    using (DataTable table = worksheet.Cells.ExportDataTable(
                                                        0,
                                                        0,
                                                        worksheet.Cells.MaxDataRow    + 1,
                                                        worksheet.Cells.MaxDataColumn + 1,
                                                        this.globalExportTableOptionsOpts))
                                                    {
                                                        foreach (DataRow row in table.Select())
                                                        {
                                                            dicProduct.TryAdd(
                                                                this.ulti.ObjectToString(row["ProductCode"]),
                                                                new Product
                                                                    {
                                                                        ProductCode = this.ulti.ObjectToString(row["ProductCode"]),
                                                                        ProductName = this.ulti.ObjectToString(row["ProductName"])
                                                                    });
                                                        }
                                                    }
                                                }
                                            }),

                                        // Customers
                                        new Task(
                                            delegate
                                            {
                                                if (!File.Exists($@"{this.applicationPath}\Database\Customers.xlsb")) return;

                                                using (var workbook = new Workbook(
                                                    $@"{this.applicationPath}\Database\Customers.xlsb",
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
                                                        worksheet.Cells.MaxDataRow    + 1,
                                                        worksheet.Cells.MaxDataColumn + 1,
                                                        this.globalExportTableOptionsOpts))
                                                    {
                                                        foreach (DataRow row in table.Select())
                                                        {
                                                            dicCustomer.TryAdd(
                                                                this.ulti.ObjectToString(row["Key"]),
                                                                new Customer
                                                                    {
                                                                        CustomerKeyCode   = this.ulti.ObjectToString(row["Code"]),
                                                                        CustomerCode      = this.ulti.ObjectToString(row["Code"]),
                                                                        CustomerName      = this.ulti.ObjectToString(row["Name"]),
                                                                        CustomerBigRegion = this.ulti.ObjectToString(row["Region"]),
                                                                        CustomerRegion    = this.ulti.ObjectToString(row["SubRegion"]),
                                                                        Company           = this.ulti.ObjectToString(row["P&L"]),
                                                                        CustomerType      = this.ulti.ObjectToString(row["Type"])
                                                                    });
                                                        }
                                                    }
                                                }
                                            }),

                                        // Orders
                                        new Task(
                                            delegate
                                            {
                                                try
                                                {
                                                    string path = $@"{this.applicationPath}\Database\Orders.xlsb";
                                                    if (!File.Exists(path)) return;

                                                    using (var workbook = new Workbook(
                                                        path,
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
                                                            worksheet.Cells.MaxDataRow    + 1,
                                                            worksheet.Cells.MaxDataColumn + 1,
                                                            this.globalExportTableOptionsOpts))
                                                        {
                                                            foreach (DataRow row in table.Select())
                                                            {
                                                                string productCode = this.ulti.ObjectToString(row["ProductCode"]);
                                                                string cusKeyCode  = this.ulti.ObjectToString(row["CustomerKeyCode"]);

                                                                for (var colIndex = 0;
                                                                     colIndex < table.Columns.Count;
                                                                     colIndex++)
                                                                {
                                                                    using (DataColumn column =
                                                                        table.Columns[colIndex])
                                                                    {
                                                                        // First check point. Is it a valid date?
                                                                        // ReSharper disable once PossibleInvalidOperationException
                                                                        // Because I'm confident about that.
                                                                        // ... it's my fucking database.
                                                                        DateTime? dateFc = this.ulti.StringToDate(
                                                                            column.ColumnName);
                                                                        if (dateFc == null) continue;

                                                                        // Second check point. Is it a valid forecast value?
                                                                        double value = this.ulti.ObjectToDouble(row[colIndex]);
                                                                        if (value <= 0) continue;

                                                                        // dicOldPo.Add(
                                                                        // ((DateTime) dateFc, productCode, cusKeyCode),
                                                                        // (new CustomerOrder
                                                                        // {
                                                                        // CustomerKeyCode = cusKeyCode,
                                                                        // QuantityOrder = poValue
                                                                        // }, false));
                                                                        listOldPo.Add(
                                                                            (((DateTime) dateFc, productCode, cusKeyCode),
                                                                            (new CustomerOrder
                                                                                 {
                                                                                     CustomerKeyCode = cusKeyCode,
                                                                                     QuantityOrder   = value
                                                                                 },
                                                                            false)));
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                                catch (Exception ex)
                                                {
                                                    this.WriteToRichTextBoxOutput(ex.Message);
                                                    throw;
                                                }
                                            })
                                    };

                // Here we go.
                Parallel.ForEach(
                    readTasks,
                    new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount },
                    task => { task.Start(); });

                await Task.WhenAll(readTasks).ConfigureAwait(false);

                // ReSharper restore HeapView.DelegateAllocation
                // ReSharper restore ImplicitlyCapturedClosure
                var listDt = new List<DataTable>();

                this.WriteToRichTextBoxOutput("Bắt đầu đọc Đơn hàng mới.", 1);

                this.TryClear();

                IOrderedEnumerable<FileInfo> files =
                    from file in new DirectoryInfo($@"{this.applicationPath}\Data\PO").GetFiles()
                    orderby file.Length descending
                    select file;

                Parallel.ForEach(
                    files,
                    new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount },
                    fileInfo =>
                    {
                        try
                        {
                            var stopwatch = new Stopwatch();
                            stopwatch.Start();

                            using (var workbook = new Workbook(
                                fileInfo.FullName,
                                new LoadOptions { MemorySetting = MemorySetting.MemoryPreference }))
                            {
                                Worksheet worksheet = workbook.Worksheets[0];
                                foreach (Worksheet ws in workbook.Worksheets)
                                {
                                    if (ws.Cells.MaxDataRow > worksheet.Cells.MaxDataRow) worksheet = ws;
                                }

                                var rowIndex = 0;
                                var colIndex = 0;

                                // Initialize First value coz of While-loop.
                                string value = worksheet.Cells[rowIndex, colIndex].Value?.ToString().Trim();

                                // Search for the very first row.
                                while (value != "VE Code" && value != "Mã Planning" && rowIndex <= 100 && colIndex <= 100)
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

                                    value = worksheet.Cells[rowIndex, colIndex].Value?.ToString().Trim();
                                }

                                // Core.
                                // Principle: Read all at once first.
                                // Then work on data later.
                                using (DataTable table = worksheet.Cells.ExportDataTable(
                                    rowIndex,
                                    colIndex,
                                    worksheet.Cells.MaxDataRow    + 1,
                                    worksheet.Cells.MaxDataColumn + 1,
                                    this.globalExportTableOptionsOpts))
                                {
                                    table.TableName = Path.GetFileNameWithoutExtension(fileInfo.Name);

                                    // To deal with the uhm, Templates having different Headers.
                                    // Please shoot me.
                                    if (!table.Columns.Contains("VE Code")) table.Columns[colIndex].ColumnName = "VE Code";

                                    // ReSharper disable once SuggestVarOrType_SimpleTypes
                                    foreach (var key in new (string oldName, string newName)[]
                                                            {
                                                                ("Tỉnh tiêu thụ", "Region"),
                                                                ("Store Code", "StoreCode"),
                                                                ("Store Name", "StoreName"),
                                                                ("Store Type", "StoreType"),
                                                                ("VE Code", "PCODE"),
                                                                ("VE Name", "PNAME")
                                                            })
                                    {
                                        if (table.Columns.Contains(key.oldName)) table.Columns[key.oldName].ColumnName = key.newName;
                                    }

                                    listDt.Add(table);
                                }
                            }

                            stopwatch.Stop();
                            this.WriteToRichTextBoxOutput(
                                $"{fileInfo.Name} - Xong trong {Math.Round(stopwatch.Elapsed.TotalSeconds, 2).ToString(CultureInfo.InvariantCulture)}s!",
                                2); // + " - Done!");
                        }
                        catch (Exception ex)
                        {
                            this.WriteToRichTextBoxOutput(ex.Message);
                            throw;
                        }
                    });

                this.WriteToRichTextBoxOutput("Bắt đầu xử lý Đơn hàng.", 1);

                this.TryClear();

                // Here comes the data handling.
                Parallel.ForEach(
                    listDt,
                    new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount },
                    table =>
                    {
                        try
                        {
                            // In an atempt to avoid missing columns.
                            // Seriously.
                            foreach (string columnName in new[] { "P&L" })
                            {
                                if (!table.Columns.Contains(columnName)) table.Columns.Add(columnName);
                            }

                            // Row layer.
                            foreach (DataRow row in table.Select())
                            {
                                // Idk why this is a thing.
                                if (string.IsNullOrEmpty(this.ulti.ObjectToString(row["PCODE"]))) continue;

                                string company                             = this.ulti.ObjectToString(row["P&L"]);
                                if (string.IsNullOrEmpty(company)) company = "VCM";

                                // Less conversion.
                                string cusKeyCode = this.ulti.GetString(
                                    $"{this.ulti.ObjectToString(row["StoreCode"])} | {this.ulti.ObjectToString(row["StoreType"])} | {company}");

                                string pCode = string.Intern(this.ulti.ObjectToString(row["PCODE"]));

                                // Product information.
                                Product product = dicProduct.GetOrAdd(
                                    pCode,
                                    new Product
                                        {
                                            ProductCode = pCode,
                                            ProductName = this.ulti.ObjectToString(row["PNAME"])
                                        });

                                // Quality of life. Get the pseudo 'best' Product Name.
                                if (string.CompareOrdinal(
                                        product.ProductName,
                                        this.ulti.ObjectToString(row["PNAME"])) <
                                    0)
                                    product.ProductName = this.ulti.ObjectToString(row["PNAME"]);

                                // Optimization, dealing with region.
                                string region = string.Intern(table.TableName.Substring(0, 2));

                                // Customer information.
                                Customer customer = dicCustomer.GetOrAdd(
                                    cusKeyCode,
                                    new Customer
                                        {
                                            CustomerKeyCode   = cusKeyCode,
                                            CustomerBigRegion = region,
                                            CustomerRegion    = this.ulti.ObjectToString(row["Region"]),
                                            CustomerCode      = this.ulti.ObjectToString(row["StoreCode"]),
                                            CustomerName      = this.ulti.ObjectToString(row["StoreName"]),
                                            CustomerType      = this.ulti.ObjectToString(row["StoreType"]),
                                            Company           = string.IsNullOrEmpty(this.ulti.ObjectToString(row["P&L"]))
                                                                    ? "VCM"
                                                                    : this.ulti.ObjectToString(row["P&L"])
                                        });

                                // Meh.
                                if (string.CompareOrdinal(
                                        customer.CustomerName,
                                        this.ulti.ObjectToString(row["StoreName"])) <
                                    0)
                                    customer.CustomerName = this.ulti.ObjectToString(row["StoreName"]);

                                // Column layer.
                                for (var colIndex = 0; colIndex < table.Columns.Count; colIndex++)
                                {
                                    using (DataColumn column = table.Columns[colIndex])
                                    {
                                        // First check point. Is it a valid date?
                                        DateTime? datePo = this.ulti.StringToDate(column.ColumnName);
                                        if (datePo == null) continue;

                                        // Second check point. Is it a valid forecast value?
                                        double poValue = this.ulti.ObjectToDouble(row[colIndex]);
                                        if (poValue <= 0) continue;

                                        // CustomerOrder order = dicPo.AddOrUpdate(
                                        // ((DateTime) datePo, pCode, cusKeyCode), (new CustomerOrder
                                        // {
                                        // //Company = customer.Company,
                                        // CustomerKeyCode = cusKeyCode,
                                        // CustomerCode = customer.CustomerCode
                                        // }, false));
                                        dicPo.AddOrUpdate(
                                            ((DateTime) datePo, pCode, cusKeyCode),
                                            (new CustomerOrder
                                                 {
                                                     CustomerCode  = customer.CustomerCode,
                                                     QuantityOrder = poValue
                                                 }, false),
                                            (key, oldValue) =>
                                                (new CustomerOrder
                                                     {
                                                         CustomerKeyCode = cusKeyCode,
                                                         CustomerCode    = customer.CustomerCode,
                                                         QuantityOrder   = oldValue.Order.QuantityOrder + poValue
                                                     }, false));

                                        // lock (myLock)
                                        // order.QuantityOrder += poValue;

                                        // lock (order)
                                        // {
                                        // order.QuantityOrder += poValue;
                                        // //order.QuantityOrderKg += poValue;
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
                    $"Xử lý xong Đơn hàng, mất: {Math.Round(watch.Elapsed.TotalSeconds, 2).ToString(CultureInfo.InvariantCulture)}s!",
                    2);

                this.TryClear();

                // ReSharper disable ImplicitlyCapturedClosure
                // ReSharper disable HeapView.DelegateAllocation
                var writeTasks = new[]
                                     {
                                         // Orders
                                         new Task(
                                             delegate
                                             {
                                                 try
                                                 {
                                                     var dicCol    = new Dictionary<string, int>();
                                                     var listTypes = new List<Type>();

                                                     // List of columns.
                                                     foreach ((string Name, Type Type) key in new (string Name, Type Type)[]
                                                                                                  {
                                                                                                      ("Region", typeof(string)),
                                                                                                      ("ProductCode", typeof(string)),
                                                                                                      ("CustomerKeyCode", typeof(string)),
                                                                                                      ("StoreCode", typeof(string)),
                                                                                                      ("StoreType", typeof(string)),
                                                                                                      ("StoreName", typeof(string)),
                                                                                                      ("SubRegion", typeof(string)),
                                                                                                      ("P&L", typeof(string))
                                                                                                  })
                                                     {
                                                         dicCol.Add(key.Name, dicCol.Count);
                                                         listTypes.Add(key.Type);
                                                     }

                                                     var listDatePo = new List<DateTime>();

                                                     // Count DateFc.
                                                     // ReSharper disable once SuggestVarOrType_SimpleTypes
                                                     foreach (var key in dicPo.Keys)
                                                     {
                                                         if (!listDatePo.Contains(key.DatePo))
                                                             listDatePo.Add(key.DatePo);
                                                     }

                                                     // ... and then add the same amount of columns.
                                                     foreach (DateTime datePo in from dateFc in listDatePo
                                                                                 orderby dateFc
                                                                                 select dateFc)
                                                     {
                                                         // Also remove all old items.
                                                         // foreach (((DateTime dateOldPo, string ProductCode, string CustomerKeyCode) key, (CustomerOrder _, bool)) oldPo in listOldPo.ToList())
                                                         // {
                                                         // if (oldPo.key.dateOldPo == datePo)
                                                         // listOldPo.Remove(oldPo);
                                                         // }
                                                         listOldPo.RemoveAll(x => x.Key.datePo == datePo);

                                                         // Add into column list.
                                                         dicCol.Add(this.ulti.DateToString(datePo, "dd-MMM-yyyy"), dicCol.Count);

                                                         // ... and type list.
                                                         listTypes.Add(typeof(double));
                                                     }

                                                     // Dictionary of rowIndex.
                                                     var dicRow = new Dictionary<string, int>(
                                                         dicProduct.Count,
                                                         StringComparer.OrdinalIgnoreCase);

                                                     // Hour of truth.
                                                     listOldPo.AddRange(dicPo.Keys.Select(key => (key, dicPo[key])));

                                                     var rowIndex = 0;

                                                     foreach (
                                                         (DateTime _, string productCode, string customerKeyCode) in
                                                         from po in listOldPo
                                                         orderby
                                                             po.Key.ProductCode,
                                                             po.Key.CustomerKeyCode
                                                         select po.Key)
                                                     {
                                                         string rowKey = $"{productCode}{customerKeyCode}";
                                                         if (dicRow.ContainsKey(rowKey)) continue;

                                                         dicRow.Add(rowKey, rowIndex);
                                                         rowIndex++;
                                                     }

                                                     var orders = new object[dicRow.Count, dicCol.Count];

                                                     // Here we go.
                                                     foreach (((DateTime _, string productCode, string customerKeyCode),
                                                         (CustomerOrder _, bool _)) in listOldPo)
                                                     {
                                                         string rowKey = $"{productCode}{customerKeyCode}";

                                                         // ("Region", typeof(string)),
                                                         // ("ProductCode", typeof(string)),
                                                         // ("CustomerKeyCode", typeof(string)),
                                                         // ("StoreCode", typeof(string)),
                                                         // ("StoreType", typeof(string)),
                                                         // ("StoreName", typeof(string)),
                                                         // ("SubRegion", typeof(string)),
                                                         // ("P&L", typeof(string)),
                                                         Customer customer = dicCustomer[customerKeyCode];

                                                         orders[dicRow[rowKey], 0] = customer.CustomerBigRegion;
                                                         orders[dicRow[rowKey], 1] = productCode;
                                                         orders[dicRow[rowKey], 2] = customerKeyCode;
                                                         orders[dicRow[rowKey], 3] = customer.CustomerCode;
                                                         orders[dicRow[rowKey], 4] = customer.CustomerType;
                                                         orders[dicRow[rowKey], 5] = customer.CustomerName;
                                                         orders[dicRow[rowKey], 6] = customer.CustomerRegion;
                                                         orders[dicRow[rowKey], 7] = customer.CustomerRegion;
                                                     }

                                                     Parallel.ForEach(
                                                         listOldPo,
                                                         new ParallelOptions
                                                             {
                                                                 MaxDegreeOfParallelism =
                                                                     Environment.ProcessorCount
                                                             },
                                                         po =>
                                                         {
                                                             try
                                                             {
                                                                 (DateTime datePo, string productCode, string customerKeyCode) = po.Key;

                                                                 string        rowKey = $"{productCode}{customerKeyCode}";
                                                                 CustomerOrder order  = po.Value.Order;

                                                                 orders[dicRow[rowKey],
                                                                        dicCol[
                                                                            this.ulti.DateToString(
                                                                                datePo,
                                                                                "dd-MMM-yyyy")]] = this.ulti.DoubleToObject(
                                                                     order.QuantityOrder);
                                                             }
                                                             catch (Exception ex)
                                                             {
                                                                 this.WriteToRichTextBoxOutput(ex.Message);
                                                                 throw;
                                                             }
                                                         });

                                                     string path = $@"{this.applicationPath}\Database\Orders.xlsx";

                                                     // ReSharper disable ArgumentsStyleOther
                                                     // ReSharper disable ArgumentsStyleNamedExpression
                                                     // ReSharper disable ArgumentsStyleStringLiteral
                                                     // ReSharper disable once ArgumentsStyleLiteral
                                                     this.ulti.ExportXmlArray((path, "Orders", new[] { orders }, dicCol.Keys.ToList(), listTypes, true, false));
//                                                                                 filePath: path,
//                                                                                 theName: "Orders",
//                                                                                 listArrays: new[] { orders },
//                                                                                 listColumnNames: dicCol.Keys.ToList(),
//                                                                                 listTypes: listTypes,
//                                                                                 yesHeader: true);
                                                     // ReSharper restore ArgumentsStyleStringLiteral
                                                     // ReSharper restore ArgumentsStyleNamedExpression
                                                     // ReSharper restore ArgumentsStyleOther

                                                     // this.ulti.ConvertExcelTypeInterop(path, "xlsx", "xlsb");
                                                     this.ulti.ConvertExcelTypeAspose(path, "xlsb");
                                                     this.ulti.DeleteEvaluationSheetInterop(path.Replace("xlsx", "xlsb"));

                                                     //// _ulti.LargeExportOneWorkbook(path, new List<DataTable> { table }, true, true);
                                                     // this.ulti.ConvertExcelTypeInterop(
                                                     //    path,
                                                     //    "xlsx",
                                                     //    "xlsb"); // Otherwise it's super fucking hard to open the file.
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
                                                     {
                                                         table.Columns.Add(key.colName, key.colType);
                                                     }

                                                     foreach (Product product in from value in dicProduct.Values
                                                                                 orderby value.ProductCode
                                                                                 select value)
                                                     {
                                                         DataRow row = table.NewRow();

                                                         row["ProductCode"] = product.ProductCode;
                                                         row["ProductName"] = product.ProductName;

                                                         table.Rows.Add(row);
                                                     }

                                                     string path = $@"{this.applicationPath}\Database\{table.TableName}.xlsx";
                                                     this.ulti.LargeExportOneWorkbook(
                                                         path,
                                                         new List<DataTable> { table },
                                                         true,
                                                         true);

                                                     // this.ulti.ConvertExcelTypeInterop(path, "xlsx", "xlsb");
                                                     this.ulti.ConvertExcelTypeAspose(path, "xlsb");
                                                     this.ulti.DeleteEvaluationSheetInterop(path.Replace("xlsx", "xlsb"));
                                                 }
                                             }),

                                         // Customers
                                         new Task(
                                             delegate
                                             {
                                                 using (var table = new DataTable { TableName = "Customers" })
                                                 {
                                                     // ReSharper disable once SuggestVarOrType_SimpleTypes
                                                     foreach (var key in new (string colName, Type colType)[]
                                                                             {
                                                                                 ("Code", typeof(string)),
                                                                                 ("Name", typeof(string)),
                                                                                 ("SubRegion", typeof(string)),
                                                                                 ("Region", typeof(string)),
                                                                                 ("Type", typeof(string)),
                                                                                 ("P&L", typeof(string)),
                                                                                 ("Key", typeof(string))
                                                                             })
                                                     {
                                                         table.Columns.Add(
                                                             key.colName,
                                                             key.colType);
                                                     }

                                                     foreach (Customer customer in
                                                         from value in dicCustomer.Values
                                                         orderby value.CustomerCode
                                                         select value)
                                                     {
                                                         DataRow row = table.NewRow();

                                                         row["Code"]      = customer.CustomerCode;
                                                         row["Name"]      = customer.CustomerName;
                                                         row["SubRegion"] = customer.CustomerRegion;
                                                         row["Region"]    = customer.CustomerBigRegion;
                                                         row["Type"]      = customer.CustomerType;
                                                         row["P&L"]       = customer.Company;
                                                         row["Key"]       = customer.CustomerKeyCode;

                                                         table.Rows.Add(row);
                                                     }

                                                     string path = $@"{this.applicationPath}\Database\{table.TableName}.xlsx";
                                                     this.ulti.LargeExportOneWorkbook(
                                                         path,
                                                         new List<DataTable> { table },
                                                         true,
                                                         true);

                                                     // this.ulti.ConvertExcelTypeInterop(path, "xlsx", "xlsb");
                                                     this.ulti.ConvertExcelTypeAspose(path, "xlsb");
                                                     this.ulti.DeleteEvaluationSheetInterop(path.Replace("xlsx", "xlsb"));
                                                 }
                                             })
                                     };

                // Here we go.
                Parallel.ForEach(
                    writeTasks,
                    new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount },
                    task => { task.Start(); });

                // Making sure every Tasks finished before proceeding.
                await Task.WhenAll(writeTasks).ConfigureAwait(false);

                // ReSharper restore HeapView.DelegateAllocation
                // ReSharper restore ImplicitlyCapturedClosure
                dicPo.Clear();
                dicProduct.Clear();
                dicCustomer.Clear();

                // The final flag.
                watch.Stop();
                this.WriteToRichTextBoxOutput(
                    $"Đã ghi vào cơ sở dữ liệu. Tổng thời gian chạy: {Math.Round(watch.Elapsed.TotalSeconds, 2).ToString(CultureInfo.InvariantCulture)}s!",
                    1);
            }
            catch (Exception ex)
            {
                this.WriteToRichTextBoxOutput(ex.Message);
                throw;
            }
            finally
            {
                this.isBackgroundworkerIdle = true;
                this.TryClear();
            }
        }
    }
}