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
        private async void ReadPurchaseOrder(object sender, DoWorkEventArgs e)
        {
            try
            {
                var watch = new Stopwatch();
                watch.Start();

                #region Initializing variables

                var dicProduct = new ConcurrentDictionary<string, Product>();
                var dicCustomer = new ConcurrentDictionary<string, Customer>();
                var dicPo =
                    new ConcurrentDictionary<(DateTime DatePo, string ProductCode, string CustomerKeyCode),
                        (CustomerOrder Order, bool)>();
                //var dicOldPo =
                //    new Dictionary<(DateTime DateFc, string ProductCode, string CustomerKeyCode),
                //        (CustomerOrder Order, bool)>();

                // Todo - Implement this instead of a Dictionary because resizing is being a bitch.
                var listOldPo
                    = new List<((DateTime datePo, string ProductCode, string CustomerKeyCode) Key,
                        (CustomerOrder Order, bool) Value)>();

                #endregion

                WriteToRichTextBoxOutput("Đọc Đơn hàng cũ từ cơ sở dữ liệu.", 1);

                #region Reading old data.

                // ReSharper disable ImplicitlyCapturedClosure
                // ReSharper disable HeapView.DelegateAllocation

                var readTasks = new[]
                {
                    // Products
                    new Task(delegate
                    {
                        if (!File.Exists($@"{_applicationPath}\Database\Products.xlsb")) return;

                        using (var xlWb = new Workbook($@"{_applicationPath}\Database\Products.xlsb",
                            new LoadOptions {MemorySetting = MemorySetting.MemoryPreference}))
                        {
                            Worksheet xlWs = xlWb.Worksheets[0];
                            using (DataTable table = xlWs.Cells.ExportDataTable(0, 0, xlWs.Cells.MaxDataRow + 1,
                                xlWs.Cells.MaxDataColumn + 1, _globalExportTableOptionsopts))
                            {
                                foreach (DataRow row in table.Select())
                                    dicProduct.TryAdd(_ulti.ObjectToString(row["ProductCode"]), new Product
                                    {
                                        ProductCode = _ulti.ObjectToString(row["ProductCode"]),
                                        ProductName = _ulti.ObjectToString(row["ProductName"])
                                    });
                            }
                        }
                    }),

                    // Customers
                    new Task(delegate
                    {
                        if (!File.Exists($@"{_applicationPath}\Database\Customers.xlsb")) return;

                        using (var xlWb = new Workbook($@"{_applicationPath}\Database\Customers.xlsb",
                            new LoadOptions {MemorySetting = MemorySetting.MemoryPreference}))
                        {
                            Worksheet xlWs = xlWb.Worksheets[0];
                            using (DataTable table = xlWs.Cells.ExportDataTable(0, 0, xlWs.Cells.MaxDataRow + 1,
                                xlWs.Cells.MaxDataColumn + 1, _globalExportTableOptionsopts))
                            {
                                foreach (DataRow row in table.Select())
                                    dicCustomer.TryAdd(_ulti.ObjectToString(row["Key"]), new Customer
                                    {
                                        CustomerKeyCode = _ulti.ObjectToString(row["Code"]),
                                        CustomerCode = _ulti.ObjectToString(row["Code"]),
                                        CustomerName = _ulti.ObjectToString(row["Name"]),
                                        CustomerBigRegion = _ulti.ObjectToString(row["Region"]),
                                        CustomerRegion = _ulti.ObjectToString(row["SubRegion"]),
                                        Company = _ulti.ObjectToString(row["P&L"]),
                                        CustomerType = _ulti.ObjectToString(row["Type"])
                                    });
                            }
                        }
                    }),

                    // Orders
                    new Task(delegate
                {
                    try
                    {
                        string path = $@"{_applicationPath}\Database\Orders.xlsb";
                        if (!File.Exists(path)) return;

                        using (var xlWb = new Workbook(path,
                            new LoadOptions {MemorySetting = MemorySetting.MemoryPreference}))
                        {
                            Worksheet xlWs = xlWb.Worksheets[0];
                            using (DataTable table = xlWs.Cells.ExportDataTable(0, 0, xlWs.Cells.MaxDataRow + 1,
                                xlWs.Cells.MaxDataColumn + 1, _globalExportTableOptionsopts))
                            {
                                foreach (DataRow row in table.Select())
                                {
                                    string productCode = _ulti.ObjectToString(row["ProductCode"]);
                                    string cusKeyCode = _ulti.ObjectToString(row["CustomerKeyCode"]);

                                    for (var colIndex = 0; colIndex < table.Columns.Count; colIndex++)
                                        using (DataColumn column = table.Columns[colIndex])
                                        {
                                            // First check point. Is it a valid date?
                                            // ReSharper disable once PossibleInvalidOperationException
                                            // Because I'm confident about that.
                                            // ... it's my fucking database.
                                            DateTime? dateFc = _ulti.StringToDate(column.ColumnName);
                                            if (dateFc == null) continue;

                                            // Second check point. Is it a valid forecast value?
                                            double poValue = _ulti.ObjectToDouble(row[colIndex]);
                                            if (poValue <= 0) continue;

                                            //dicOldPo.Add(
                                            //    ((DateTime) dateFc, productCode, cusKeyCode),
                                            //    (new CustomerOrder
                                            //    {
                                            //        CustomerKeyCode = cusKeyCode,
                                            //        QuantityOrder = poValue
                                            //    }, false));

                                            listOldPo.Add(
                                                (((DateTime) dateFc, productCode, cusKeyCode),
                                                (new CustomerOrder
                                                {
                                                    CustomerKeyCode = cusKeyCode,
                                                    QuantityOrder = poValue
                                                }, false)));
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
                })
                };

                // Here we go.
                Parallel.ForEach(readTasks, new ParallelOptions {MaxDegreeOfParallelism = Environment.ProcessorCount},
                    task => { task.Start(); });

                await Task.WhenAll(readTasks);

                // ReSharper restore HeapView.DelegateAllocation
                // ReSharper restore ImplicitlyCapturedClosure

                #endregion

                var listDt = new List<DataTable>();

                WriteToRichTextBoxOutput("Bắt đầu đọc Đơn hàng mới.", 1);

                TryClear();

                #region Reading new data.

                Parallel.ForEach(new DirectoryInfo($@"{_applicationPath}\Data\PO").GetFiles(),
                    new ParallelOptions {MaxDegreeOfParallelism = Environment.ProcessorCount},
                    fileInfo =>
                    {
                        try
                        {
                            var stopwatch = new Stopwatch();
                            stopwatch.Start();

                            using (var xlWb = new Workbook(fileInfo.FullName,
                                new LoadOptions {MemorySetting = MemorySetting.MemoryPreference}))
                            {
                                Worksheet xlWs = xlWb.Worksheets[0];
                                foreach (Worksheet ws in xlWb.Worksheets)
                                {
                                    if (ws.Cells.MaxDataRow > xlWs.Cells.MaxDataRow)
                                        xlWs = ws;
                                }

                                var rowIndex = 0;
                                var colIndex = 0;

                                // Initialize First value coz of While-loop.
                                string value = xlWs.Cells[rowIndex, colIndex].Value?.ToString().Trim();

                                // Search for the very first row.
                                while (value != "VE Code" && value != "Mã Planning" && rowIndex <= 100 &&
                                       colIndex <= 100)
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

                                    if (!table.Columns.Contains("VE Code"))
                                        table.Columns[colIndex].ColumnName = "VE Code";

                                    foreach ((string oldName, string newName) in new[]
                                    {
                                        ("Tỉnh tiêu thụ", "Region"),
                                        ("Store Code", "StoreCode"),
                                        ("Store Name", "StoreName"),
                                        ("Store Type", "StoreType"),
                                        ("VE Code", "PCODE"),
                                        ("VE Name", "PNAME"),
                                    })
                                        if (table.Columns.Contains(oldName))
                                            table.Columns[oldName].ColumnName = newName;

                                    listDt.Add(table);
                                }
                            }

                            stopwatch.Stop();
                            WriteToRichTextBoxOutput(
                                message:
                                $"{fileInfo.Name} - Xong trong {Math.Round(stopwatch.Elapsed.TotalSeconds, 2).ToString(CultureInfo.InvariantCulture)}s!",
                                importanceLevel: 2); // + " - Done!");
                        }
                        catch (Exception ex)
                        {
                            WriteToRichTextBoxOutput(ex.Message);
                            throw;
                        }
                    });

                #endregion

                WriteToRichTextBoxOutput("Bắt đầu xử lý Đơn hàng.", 1);

                TryClear();

                #region Handling Data.

                // Here comes the data handling.
                Parallel.ForEach(listDt, new ParallelOptions {MaxDegreeOfParallelism = Environment.ProcessorCount},
                    table =>
                    {
                        try
                        {
                            // Row layer.
                            foreach (DataRow row in table.Select())
                            {
                                // Idk why this is a thing.
                                if (string.IsNullOrEmpty(_ulti.ObjectToString(row["PCODE"]))) continue;

                                // Less conversion.
                                string cusKeyCode = _ulti.GetString(
                                    $"{_ulti.ObjectToString(row["StoreCode"])} | {_ulti.ObjectToString(row["StoreType"])} | {_ulti.ObjectToString(row["P&L"])}");

                                string pCode = string.Intern(_ulti.ObjectToString(row["PCODE"]));

                                // Product information.
                                Product product = dicProduct.GetOrAdd(pCode, new Product
                                {
                                    ProductCode = pCode,
                                    ProductName = _ulti.ObjectToString(row["PNAME"])
                                });

                                // Quality of life. Get the pseudo 'best' Product Name.
                                if (string.CompareOrdinal(product.ProductName, _ulti.ObjectToString(row["PNAME"])) < 0)
                                    product.ProductName = _ulti.ObjectToString(row["PNAME"]);

                                // Optimization, dealing with region.
                                string region = string.Intern(table.TableName.Substring(0, 2));

                                // Customer information.
                                Customer customer = dicCustomer.GetOrAdd(cusKeyCode, new Customer
                                {
                                    CustomerBigRegion = region,
                                    CustomerRegion = _ulti.ObjectToString(row["Region"]),
                                    Company = _ulti.ObjectToString(row["P&L"]),
                                    CustomerKeyCode = cusKeyCode,
                                    CustomerCode = _ulti.ObjectToString(row["StoreCode"]),
                                    CustomerName = _ulti.ObjectToString(row["StoreName"]),
                                    CustomerType = _ulti.ObjectToString(row["StoreType"])
                                });

                                // Meh.
                                if (string.CompareOrdinal(customer.CustomerName,
                                        _ulti.ObjectToString(row["StoreName"])) < 0)
                                    customer.CustomerName = _ulti.ObjectToString(row["StoreName"]);

                                // Column layer.
                                for (var colIndex = 0; colIndex < table.Columns.Count; colIndex++)
                                    using (DataColumn column = table.Columns[colIndex])
                                    {
                                        // First check point. Is it a valid date?
                                        DateTime? datePo = _ulti.StringToDate(column.ColumnName);
                                        if (datePo == null) continue;

                                        // Second check point. Is it a valid forecast value?
                                        double poValue = _ulti.ObjectToDouble(row[colIndex]);
                                        if (poValue <= 0) continue;

                                        //CustomerOrder order = dicPo.AddOrUpdate(
                                        //    ((DateTime) datePo, pCode, cusKeyCode), (new CustomerOrder
                                        //    {
                                        //        //Company = customer.Company,
                                        //        CustomerKeyCode = cusKeyCode,
                                        //        CustomerCode = customer.CustomerCode
                                        //    }, false));

                                        dicPo.AddOrUpdate(((DateTime) datePo, pCode, cusKeyCode), (new CustomerOrder
                                        {
                                            //Company = customer.Company,
                                            CustomerKeyCode = cusKeyCode,
                                            CustomerCode = customer.CustomerCode,
                                            QuantityOrder = poValue
                                        }, false), (key, oldValue) => (new CustomerOrder
                                        {
                                            CustomerKeyCode = cusKeyCode,
                                            CustomerCode = customer.CustomerCode,
                                            QuantityOrder = oldValue.Order.QuantityOrder + poValue
                                        }, false));

                                        //lock (myLock)
                                        //    order.QuantityOrder += poValue;

                                        //lock (order)
                                        //{
                                        //    order.QuantityOrder += poValue;
                                        //    //order.QuantityOrderKg += poValue;
                                        //}
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

                WriteToRichTextBoxOutput(
                    $"Xử lý xong Đơn hàng, mất: {Math.Round(watch.Elapsed.TotalSeconds, 2).ToString(CultureInfo.InvariantCulture)}s!",
                    2);

                TryClear();

                #region Write down Data.

                // ReSharper disable ImplicitlyCapturedClosure
                // ReSharper disable HeapView.DelegateAllocation

                var writeTasks = new[]
                {
                    // Orders
                    new Task(delegate
                    {
                        try
                        {
                            var listColumns = new List<string>();
                            var listTypes = new List<Type>();

                            foreach ((string colName, Type colType) in new[]
                            {
                                ("ProductCode", typeof(string)),
                                ("CustomerKeyCode", typeof(string))
                            })
                            {
                                listColumns.Add(colName);
                                listTypes.Add(colType);
                            }

                            var listDatePo = new List<DateTime>();

                            // Count DateFc.
                            foreach ((DateTime datePo, string _, string _) in dicPo.Keys)
                                if (!listDatePo.Contains(datePo))
                                    listDatePo.Add(datePo);

                            // ... and then add the same amount of columns.
                            foreach (DateTime datePo in
                                from dateFc in listDatePo
                                orderby dateFc
                                select dateFc)
                            {
                                // Also remove all old items.
                                //foreach (((DateTime dateOldPo, string ProductCode, string CustomerKeyCode) key, (CustomerOrder _, bool)) oldPo in listOldPo.ToList())
                                //{
                                //    if (oldPo.key.dateOldPo == datePo)
                                //        listOldPo.Remove(oldPo);
                                //}

                                listOldPo.RemoveAll(x => x.Key.datePo == datePo);
                                listColumns.Add(_ulti.DateToString(datePo, "dd-MMM-yyyy"));
                                listTypes.Add(typeof(double));
                            }

                            // Dictionary of rowIndex.
                            var dicRow =
                                new Dictionary<string, int>(dicProduct.Count, StringComparer.OrdinalIgnoreCase);

                            // Hour of truth.
                            listOldPo.AddRange(dicPo.Keys.Select(key => (key, dicPo[key])));

                            var rowIndex = 0;
                            foreach ((DateTime _, string productCode, string customerKeyCode) in
                                from po in listOldPo
                                orderby po.Key.ProductCode, po.Key.CustomerKeyCode
                                select po.Key)
                            {
                                string rowKey = $"{productCode}{customerKeyCode}";
                                if (dicRow.ContainsKey(rowKey)) continue;

                                dicRow.Add(rowKey, rowIndex);
                                rowIndex++;
                            }

                            var orders = new object[dicRow.Count, listColumns.Count];
                            foreach (((DateTime _, string ProductCode, string CustomerKeyCode) Key, (CustomerOrder _,
                                bool)) po in listOldPo)
                            {
                                string rowKey = $"{po.Key.ProductCode}{po.Key.CustomerKeyCode}";
                                orders[dicRow[rowKey], 0] = po.Key.ProductCode;
                                orders[dicRow[rowKey], 1] = po.Key.CustomerKeyCode;
                            }

                            Parallel.ForEach(listOldPo,
                                new ParallelOptions {MaxDegreeOfParallelism = Environment.ProcessorCount},
                                po =>
                                {
                                    try
                                    {
                                        (DateTime datePo, string productCode, string customerKeyCode) = po.Key;
                                        string rowKey = $"{productCode}{customerKeyCode}";
                                        CustomerOrder order = po.Value.Order;

                                        orders[dicRow[rowKey],
                                                listColumns.IndexOf(_ulti.DateToString(datePo, "dd-MMM-yyyy"))] =
                                            _ulti.DoubleToObject(order.QuantityOrder);
                                    }
                                    catch (Exception ex)
                                    {
                                        WriteToRichTextBoxOutput(ex.Message);
                                        throw;
                                    }
                                });

                            string path = $@"{_applicationPath}\Database\Orders.xlsx";
                            _ulti.ExportXmlArray(
                                filePath: path,
                                theName: "Orders",
                                listArrays: new[] {orders},
                                listColumnNames: listColumns,
                                listTypes: listTypes,
                                yesHeader: true);
                            //_ulti.LargeExportOneWorkbook(path, new List<DataTable> { table }, true, true);
                            _ulti.ConvertExcelTypeInterop(path, "xlsx",
                                "xlsb"); // Otherwise it's super fucking hard to open the file.
                        }

                        catch (Exception ex)
                        {
                            WriteToRichTextBoxOutput(ex.Message);
                            throw;
                        }
                    }),

                    // Products
                    new Task(delegate
                    {
                        using (var table = new DataTable {TableName = "Products"})
                        {
                            foreach ((string colName, Type colType) in new[]
                            {
                                ("ProductCode", typeof(string)),
                                ("ProductName", typeof(string))
                            })
                                table.Columns.Add(colName, colType);

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

                            string path = $@"{_applicationPath}\Database\{table.TableName}.xlsx";
                            _ulti.LargeExportOneWorkbook(path, new List<DataTable> {table}, true, true);
                            _ulti.ConvertExcelTypeInterop(path, "xlsx", "xlsb");
                        }
                    }),

                    // Customers
                    new Task(delegate
                    {
                        using (var table = new DataTable {TableName = "Customers"})
                        {
                            foreach ((string colName, Type colType) in new[]
                            {
                                ("Code", typeof(string)),
                                ("Name", typeof(string)),
                                ("SubRegion", typeof(string)),
                                ("Region", typeof(string)),
                                ("Type", typeof(string)),
                                ("P&L", typeof(string)),
                                ("Key", typeof(string)),

                            })
                                table.Columns.Add(colName, colType);

                            foreach (Customer customer in
                                from value in dicCustomer.Values
                                orderby value.CustomerCode
                                select value)
                            {
                                DataRow row = table.NewRow();

                                row["Code"] = customer.CustomerCode;
                                row["Name"] = customer.CustomerName;
                                row["SubRegion"] = customer.CustomerRegion;
                                row["Region"] = customer.CustomerBigRegion;
                                row["Type"] = customer.CustomerType;
                                row["P&L"] = customer.Company;
                                row["Key"] = customer.CustomerKeyCode;

                                table.Rows.Add(row);
                            }

                            string path = $@"{_applicationPath}\Database\{table.TableName}.xlsx";
                            _ulti.LargeExportOneWorkbook(path, new List<DataTable> {table}, true, true);
                            _ulti.ConvertExcelTypeInterop(path, "xlsx", "xlsb");
                        }
                    })
                };

                // Here we go.
                Parallel.ForEach(writeTasks, new ParallelOptions {MaxDegreeOfParallelism = Environment.ProcessorCount},
                    task => { task.Start(); });

                // Making sure every Tasks finished before proceeding.
                await Task.WhenAll(writeTasks);

                // ReSharper restore HeapView.DelegateAllocation
                // ReSharper restore ImplicitlyCapturedClosure

                #endregion

                dicPo.Clear();
                dicProduct.Clear();
                dicCustomer.Clear();

                // The final flag.
                watch.Stop();
                WriteToRichTextBoxOutput(
                    $"Đã ghi vào cơ sở dữ liệu. Tổng thời gian chạy: {Math.Round(watch.Elapsed.TotalSeconds, 2).ToString(CultureInfo.InvariantCulture)}s!",
                    1);
            }
            // Just, why?
            catch (Exception ex)
            {
                WriteToRichTextBoxOutput(ex.Message);
                throw;
            }
            finally
            {
                TryClear();
            }
        }


    }
}