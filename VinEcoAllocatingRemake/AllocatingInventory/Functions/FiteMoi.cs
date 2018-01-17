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

namespace VinEcoAllocatingRemake.AllocatingInventory
{
    public partial class AllocatingInventory
    {
        private async void FiteMoi(object sender, DoWorkEventArgs e)
        {
            try
            {
                // Plz me first u no add things before me do me jobs.
                var watch = new Stopwatch();
                watch.Start();

                DateTime dateFrom = DateTime.Today;
                DateTime dateTo = DateTime.Today;

                var distance = new Dictionary<(string supRegion, string cusRegion), int>(4)
                {
                    {("MB", "MB"), 1},
                    {("MN", "MN"), 0},
                    {("LD", "MB"), 3},
                    {("LD", "MN"), 0},
                };

                await Dispatcher.BeginInvoke((Action)(() =>
                {
                    dateFrom = DateFromCalendar.DisplayDate;
                    dateTo = DateToCalendar.DisplayDate;

                    distance = new Dictionary<(string supRegion, string cusRegion), int>(4)
                    {
                        {("MB", "MB"), int.Parse(NorthNorth.Text)},
                        {("MN", "MN"), int.Parse(SouthSouth.Text)},
                        {("LD", "MB"), int.Parse(MidNorth.Text)},
                        {("LD", "MN"), int.Parse(MidSouth.Text)},
                    };
                }));

                int maxDistance = distance.Values.Max();

                dateTo = dateTo > dateFrom ? dateTo : dateFrom;
                
                #region Initializing variables

                var products = new ConcurrentDictionary<string, Product>();
                var suppliers = new ConcurrentDictionary<string, Supplier>();
                var dicFc =
                    new Dictionary<(DateTime DateFc, string Region, string ProductCode), 
                        (SupplierForecast Supply, bool)>();
                
                var customers = new ConcurrentDictionary<string, Customer>();
                var dicPo =
                    new Dictionary<(DateTime DatePo, string Region, string ProductCode), 
                        (CustomerOrder Order, bool)>();

                var dicMoq = new Dictionary<string, double>
                {
                    {"K01901", 0.3}, // Chanh có hạt
                    {"K02201", 0.3}, // Chanh không hạt
                    {"C07101", 0.1}, // Ớt ngọt ( chuông ) đỏ
                    {"C07201", 0.1}, // Ớt ngọt ( chuông ) vàng
                    {"C07301", 0.1}, // Ớt ngọt ( chuông ) xanh
                    {"B00201", 0.3}, // Dọc mùng ( bạc hà )
                    {"C01801", 0.3}, // Cà chua cherry đỏ
                    {"C04401", 0.3}, // Đậu bắp xanh
                };

                #endregion

                WriteToRichTextBoxOutput("Bắt đầu đọc Data", 1);

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
                                    products.TryAdd(_ulti.ObjectToString(row["ProductCode"]), new Product
                                    {
                                        ProductCode = _ulti.ObjectToString(row["ProductCode"]),
                                        ProductName = _ulti.ObjectToString(row["ProductName"])
                                    });
                            }
                        }
                    }),

                    // Suppliers
                    new Task(delegate
                    {
                        if (!File.Exists($@"{_applicationPath}\Database\Suppliers.xlsb")) return;

                        using (var xlWb = new Workbook($@"{_applicationPath}\Database\Suppliers.xlsb",
                            new LoadOptions {MemorySetting = MemorySetting.MemoryPreference}))
                        {
                            Worksheet xlWs = xlWb.Worksheets[0];
                            using (DataTable table = xlWs.Cells.ExportDataTable(0, 0, xlWs.Cells.MaxDataRow + 1,
                                xlWs.Cells.MaxDataColumn + 1, _globalExportTableOptionsopts))
                            {
                                foreach (DataRow row in table.Select())
                                    suppliers.TryAdd(_ulti.ObjectToString(row["SupplierCode"]), new Supplier
                                    {
                                        SupplierRegion = _ulti.ObjectToString(row["SupplierRegion"]),
                                        SupplierType = _ulti.ObjectToString(row["SupplierType"]),
                                        SupplierCode = _ulti.ObjectToString(row["SupplierCode"]),
                                        SupplierName = _ulti.ObjectToString(row["SupplierName"])
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
                                    customers.TryAdd(_ulti.ObjectToString(row["Key"]), new Customer
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
                };

                // Here we go.
                Parallel.ForEach(readTasks, new ParallelOptions {MaxDegreeOfParallelism = Environment.ProcessorCount},
                    task => { task.Start(); });

                // Gonna wait for all reading tasks to finish.
                await Task.WhenAll(readTasks);
                WriteToRichTextBoxOutput("Đọc xong Data - Phần 1.", 2);

                readTasks = new[]
                {
                    // Forecasts
                    new Task(delegate
                    {
                        // Safeguard
                        if (!File.Exists($@"{_applicationPath}\Database\Forecasts.xlsb"))
                        {
                            WriteToRichTextBoxOutput("Không có Database Forecast.");
                            return;
                        }

                        using (var xlWb = new Workbook($@"{_applicationPath}\Database\Forecasts.xlsb",
                            new LoadOptions {MemorySetting = MemorySetting.MemoryPreference}))
                        {
                            Worksheet xlWs = xlWb.Worksheets[0];
                            using (DataTable table = xlWs.Cells.ExportDataTable(0, 0, xlWs.Cells.MaxDataRow + 1,
                                xlWs.Cells.MaxDataColumn + 1, _globalExportTableOptionsopts))
                            {
                                var colFirst = 0;
                                var colLast = 0;

                                for (var colIndex = 0; colIndex < table.Columns.Count; colIndex++)
                                {
                                    using (DataColumn column = table.Columns[colIndex])
                                    {
                                        DateTime? dateFc = _ulti.StringToDate(_ulti.GetString(column.ColumnName));
                                        if (dateFc == null) continue;

                                        if (dateFc == dateFrom.AddDays(-maxDistance))
                                            colFirst = colIndex;

                                        if (dateFc != dateTo.AddDays(maxDistance)) continue;

                                        colLast = colIndex;
                                        break;
                                    }
                                }

                                foreach (DataRow row in table.Select())
                                {
                                    string productCode = _ulti.ObjectToString(row["ProductCode"]);
                                    string supplierCode = _ulti.ObjectToString(row["SupplierCode"]);

                                    for (int colIndex = colFirst; colIndex <= colLast; colIndex++)
                                        using (DataColumn column = table.Columns[colIndex])
                                        {
                                            // First check point. Is it a valid date?
                                            DateTime? dateFc = _ulti.StringToDate(column.ColumnName);
                                            // FiteMoi specific Validation for date.
                                            //if (dateFc == null || dateFc > dateTo || dateFc < dateFrom) continue;

                                            // Second check point. Is it a valid forecast value?
                                            double fcValue = _ulti.ObjectToDouble(row[colIndex]);
                                            if (fcValue <= 0) continue;

                                            dicFc.Add(
                                                // ReSharper disable once PossibleInvalidOperationException
                                                ((DateTime) dateFc, supplierCode, productCode),
                                                (new SupplierForecast
                                                {
                                                    QualityControlPass = true,
                                                    SupplierCode = supplierCode,
                                                    FullOrder = _ulti.ObjectToInt(row["FullOrder"]) == 1,
                                                    CrossRegion = _ulti.ObjectToInt(row["CrossRegion"]) == 1,
                                                    LabelVinEco = _ulti.ObjectToInt(row["Label"]) == 1,
                                                    Level = (byte) _ulti.ObjectToInt(row["Level"])
                                                }, false));
                                        }
                                }
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
                                    var colFirst = 0;
                                    var colLast = 0;

                                    for (var colIndex = 0; colIndex < table.Columns.Count; colIndex++)
                                    {
                                        using (DataColumn column = table.Columns[colIndex])
                                        {
                                            DateTime? dateFc = _ulti.StringToDate(_ulti.GetString(column.ColumnName));
                                            if (dateFc == null) continue;

                                            if (dateFc == dateFrom.AddDays(-maxDistance))
                                                colFirst = colIndex;

                                            if (dateFc != dateTo.AddDays(maxDistance)) continue;

                                            colLast = colIndex;
                                            break;
                                        }
                                    }

                                    foreach (DataRow row in table.Select())
                                    {
                                        string productCode = _ulti.ObjectToString(row["ProductCode"]);
                                        string cusKeyCode = _ulti.ObjectToString(row["CustomerKeyCode"]);

                                        for (int colIndex = colFirst; colIndex <= colLast; colIndex++)
                                            using (DataColumn column = table.Columns[colIndex])
                                            {
                                                // First check point. Is it a valid date?
                                                // ReSharper disable once PossibleInvalidOperationException
                                                // Because I'm confident about that.
                                                // ... it's my fucking database.
                                                DateTime? datePo = _ulti.StringToDate(_ulti.GetString(column.ColumnName));
                                                //if (datePo == null || datePo > dateTo || datePo < dateFrom) continue;

                                                // Second check point. Is it a valid forecast value?
                                                double poValue = _ulti.ObjectToDouble(row[colIndex]);
                                                if (poValue <= 0) continue;

                                                dicPo.Add(
                                                    // ReSharper disable once PossibleInvalidOperationException
                                                    ((DateTime) datePo, cusKeyCode, productCode),
                                                    (new CustomerOrder
                                                    {
                                                        CustomerKeyCode = cusKeyCode,
                                                        QuantityOrder = poValue
                                                    }, false));
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

                // Gonna wait for all reading tasks to finish.
                await Task.WhenAll(readTasks);
                WriteToRichTextBoxOutput("Đọc xong Data - Phần 2.", 2);

                var listDatePo = new List<DateTime>();
                foreach ((DateTime datePo, string _, string _) in dicPo.Keys)
                {
                    if (!listDatePo.Contains(datePo))
                        listDatePo.Add(datePo);
                }

                WriteToRichTextBoxOutput("Here goes nothing.");

                foreach (string priority in new[]
                {
                    "B2B",
                    //"VM+ VinEco Priority",
                    "VM+ VinEco",
                    //"VM Priority",
                    //"VM+",
                    //"VM",
                    string.Empty
                })
                    Coord(priority);

                void Coord(string priority = "")
                {
                    try
                    {
                        var localWatch = new Stopwatch();

                        foreach (DateTime datePo in listDatePo)
                        {
                            localWatch.Restart();

                            foreach (string pCode in products.Keys)
                            {
                                //if (!listPo.Any(po =>
                                //    po.Key.ProductCode == pCode &&
                                //    (priority == string.Empty ||
                                //     //po.Value.Order.CustomerCode.IndexOf($"| {priority} |",
                                //     //    StringComparison.OrdinalIgnoreCase) >= 0)))
                                //     (po.Value.Order.CustomerCode?.Contains($"| {priority} |") ?? false))))
                                //    continue;

                                //if (listFc.All(fc => fc.Key.ProductCode != pCode))
                                //    continue;

                                Dictionary<(DateTime DatePo, string CusKeyCode, string ProductCode),
                                    (CustomerOrder Order, bool)> localPo =
                                    dicPo.Where(po =>
                                            po.Key.ProductCode == pCode &&
                                            (priority == string.Empty ||
                                             /*customers[po.Value.Order.CustomerCode].CustomerType == priority*/
                                             (po.Value.Order.CustomerCode?.Contains($"| {priority} |") ?? false)))
                                        .ToDictionary(x => x.Key, x => x.Value);

                                Dictionary<(DateTime DateFc, string SupplierCode, string ProductCode),
                                    (SupplierForecast Supply, bool)> localFc =
                                    dicFc.Where(fc => fc.Key.ProductCode == pCode)
                                        .ToDictionary(x => x.Key, x => x.Value);

                                //localFc.RemoveAll(fc => fc.Key.ProductCode != pCode);

                                if (!localFc.Any()) continue;

                                //localPo.RemoveAll(po => po.Key.ProductCode != pCode ||
                                //                        (priority != string.Empty ||
                                //                         customers[po.Value.Order.CustomerCode].CustomerType != priority));

                                // Just, skip.
                                if (!localPo.Any()) continue;

                                //var poNorth =
                                //    new List<((DateTime DatePo, string Region, string ProductCode) Key, (CustomerOrder
                                //        Order, bool) Value)>(localPo);
                                //poNorth.RemoveAll(po =>
                                //    po.Key.DatePo != datePo || 
                                //    po.Key.Region != _ulti.GetString("MB"));

                                //var poSouth =
                                //    new List<((DateTime DatePo, string Region, string ProductCode) Key, (CustomerOrder
                                //        Order, bool) Value)>(localPo);
                                //poSouth.RemoveAll(po =>
                                //    po.Key.DatePo != datePo.AddDays(-distance[("LD", "MB")] + distance[("LD", "MN")]) || 
                                //    po.Key.Region != _ulti.GetString("MN"));

                                var poNorth =
                                    new Dictionary<(DateTime DatePo, string Region, string ProductCode),
                                        (CustomerOrder Order, bool)>();

                                var poSouth =
                                    new Dictionary<(DateTime DatePo, string Region, string ProductCode),
                                        (CustomerOrder Order, bool)>();

                                double sumPoNorth = 0;
                                double sumPoSouth = 0;
                                // ReSharper disable once ForCanBeConvertedToForeach
                                foreach ((DateTime DatePo, string CusKeyCode, string ProductCode) key in localPo.Keys)
                                {
                                    (CustomerOrder Order, bool) value = localPo[key];
                                    if (key.DatePo == datePo && customers[key.CusKeyCode].CustomerBigRegion == _ulti.GetString("MB"))
                                        //if ((x.Key.DatePo, x.Key.Region, x.Key.ProductCode).Equals((datePo, "MB", pCode)))
                                    {
                                        poNorth.Add(key, value);
                                        sumPoNorth += value.Order.QuantityOrder;
                                    }
                                    else if (key.DatePo ==
                                        datePo.AddDays(-distance[("LD", "MB")] + distance[("LD", "MN")]) &&
                                        customers[key.CusKeyCode].CustomerBigRegion == "MN")
                                        //if ((x.Key.DatePo, x.Key.Region, x.Key.ProductCode).Equals(
                                        //    (datePo.AddDays(-dicDistance[("LĐ", "MB")] + dicDistance[("LĐ", "MN")]), "MN",
                                        //    pCode))) 
                                    {
                                        poSouth.Add(key, value);
                                        sumPoSouth += value.Order.QuantityOrder;
                                    }
                                }

                                //var fcNorth =
                                //    new List<((DateTime DateFc, string Region, string ProductCode) Key, (
                                //        SupplierForecast Supply, bool) Value)>(localFc);
                                //fcNorth.RemoveAll(fc =>
                                //    fc.Key.DateFc != datePo.AddDays(-distance[("MB", "MB")]) || 
                                //    fc.Key.Region != "MB");

                                //var fcMid =
                                //    new List<((DateTime DateFc, string Region, string ProductCode) Key, (
                                //        SupplierForecast Supply, bool) Value)>(localFc);
                                //fcMid.RemoveAll(fc =>
                                //    fc.Key.DateFc != datePo.AddDays(-Math.Max(distance[("LD", "MB")], distance[("LD", "MN")])) || 
                                //    fc.Key.Region != "LD");

                                //var fcSouth =
                                //    new List<((DateTime DateFc, string Region, string ProductCode) Key, (
                                //        SupplierForecast Supply, bool) Value)>(localFc);
                                //fcSouth.RemoveAll(fc =>
                                //    fc.Key.DateFc != datePo.AddDays(-distance[("MN", "MN")]) || 
                                //    fc.Key.Region != "MN");

                                var fcNorth =
                                    new Dictionary<(DateTime DateFc, string Region, string ProductCode), (
                                        SupplierForecast Supply, bool)>();

                                var fcMid =
                                    new Dictionary<(DateTime DateFc, string Region, string ProductCode), (
                                        SupplierForecast Supply, bool)>();

                                var fcSouth =
                                    new Dictionary<(DateTime DateFc, string Region, string ProductCode), (
                                        SupplierForecast Supply, bool)>();

                                double sumFcNorth = 0;
                                double sumFcMid = 0;
                                double sumFcSouth = 0;
                                // ReSharper disable once ForCanBeConvertedToForeach
                                foreach ((DateTime DateFc, string SupplierCode, string ProductCode) key in localFc.Keys)
                                {
                                    (SupplierForecast Supply, bool) value = localFc[key];
                                    //((DateTime DateFc, string Region, string ProductCode) Key,
                                    //    (SupplierForecast Supply, bool) Value) fc = listFc[index];

                                    if (key.DateFc == datePo.AddDays(distance[("MB", "MB")]) &&
                                        suppliers[key.SupplierCode].SupplierRegion == "MB")
                                    {
                                        fcNorth.Add(key, value);
                                        sumFcNorth += value.Supply.QuantityForecast;
                                    }
                                    else if (key.DateFc != datePo.AddDays(-Math.Max(distance[("LD", "MB")], distance[("LD", "MN")])) ||
                                        suppliers[key.SupplierCode].SupplierRegion != "LD")
                                    {
                                        fcMid.Add(key, value);
                                        sumFcMid += value.Supply.QuantityForecast;
                                    }
                                    else if (key.DateFc != datePo.AddDays(-distance[("MN", "MN")]) ||
                                        suppliers[key.SupplierCode].SupplierRegion != "MN")
                                    {
                                        fcSouth.Add(key, value);
                                        sumFcSouth += value.Supply.QuantityForecast;
                                    }
                                }

                                //var count = 0;
                                //for (var index = 0; index < listPo.Count; index++)
                                //{
                                //    if (dicPo.Keys.ElementAt(index).DatePo == datePo)
                                //        count++;
                                //}

                                //var poNorth = new Dictionary<(DateTime DatePo, string ProductCode, string CustomerKeyCode),
                                //    (CustomerOrder Order, bool)>(count);

                                //for (var index = 0; index < dicPo.Count; index++)
                                //{
                                //    (DateTime DatePo, string ProductCode, string CustomerKeyCode) key =
                                //        dicPo.Keys.ElementAt(index);

                                //    if (key.DatePo != datePo || dicPo[key].Item2 == false) continue;
                                //    poNorth.Add(key, dicPo[key]);
                                //}

                                // Todo - CHECK THIS.
                                //double sumPoNorth = 0;
                                //foreach (((DateTime _, string _, string _), (CustomerOrder Order, bool) value) in
                                //    poNorth) sumPoNorth += value.Order.QuantityOrder;

                                //double sumPoSouth = 0;
                                //foreach (((DateTime _, string _, string _), (CustomerOrder Order, bool) value) in
                                //    poSouth) sumPoSouth += value.Order.QuantityOrder;

                                //double sumFcNorth = 0;
                                //foreach (((DateTime _, string _, string _), (SupplierForecast Supply, bool) value) in
                                //    fcNorth) sumFcNorth += value.Supply.QuantityForecast;

                                //double sumFcMid = 0;
                                //foreach (((DateTime _, string _, string _), (SupplierForecast Supply, bool) value) in
                                //    fcMid) sumFcMid += value.Supply.QuantityForecast;

                                //double sumFcSouth = 0;
                                //foreach (((DateTime _, string _, string _), (SupplierForecast Supply, bool) value) in
                                //    fcSouth) sumFcSouth += value.Supply.QuantityForecast;

                                //WriteToRichTextBoxOutput($"Date: {_ulti.DateToString(datePo, "dd-MMM-yy")} || Product: {pCode} - {products[pCode].ProductName}");
                                //WriteToRichTextBoxOutput($"PO: North: {sumPoNorth} || South: {sumPoSouth}");
                                //WriteToRichTextBoxOutput($"FC: North: {sumFcNorth} || Mid: {sumFcMid} || South: {sumFcSouth}");
                                //WriteToRichTextBoxOutput();

                                //listPo = listPo.Except(localPo).ToList();

                                //return;
                            }
                            
                            WriteToRichTextBoxOutput($"{_ulti.DateToString(datePo, "dd-MMM-yyyy")}: {Math.Round(localWatch.Elapsed.TotalSeconds, 2).ToString(CultureInfo.InvariantCulture)}s!");
                        }

                        localWatch.Stop();
                    }
                    catch (Exception ex)
                    {
                        WriteToRichTextBoxOutput(ex.Message);
                        throw;
                    }
                }

                // The final flag.
                watch.Stop();
                WriteToRichTextBoxOutput(
                    $"Đã ghi vào cơ sở dữ liệu. Tổng thời gian chạy: {Math.Round(watch.Elapsed.TotalSeconds, 2).ToString(CultureInfo.InvariantCulture)}s!",
                    1);
            }
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