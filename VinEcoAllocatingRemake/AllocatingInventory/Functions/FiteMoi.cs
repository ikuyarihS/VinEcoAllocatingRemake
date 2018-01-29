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
        ///     Fite moi!.
        /// </summary>
        /// <param name="sender">
        ///     The sender.
        /// </param>
        /// <param name="e">
        ///     The e.
        /// </param>
        private async void FiteMoi(
            object sender,
            DoWorkEventArgs e)
        {
            try
            {
                // Plz me first u no add things before me do me jobs.
                var watch = new Stopwatch();
                watch.Start();

                DateTime dateFrom = DateTime.Today;
                DateTime dateTo = DateTime.Today;

                var distance =
                    new Dictionary<(string supRegion, string cusRegion), int>(4)
                    {
                        {("MB", "MB"), 1},
                        {("MN", "MN"), 0},
                        {("LD", "MB"), 3},
                        {("LD", "MN"), 0}
                    };

                // ReSharper disable once AsyncConverter.ConfigureAwaitHighlighting
                await Dispatcher.BeginInvoke(
                    (Action) (() =>
                    {
                        dateFrom = DateFromCalendar.DisplayDate;
                        dateTo = DateToCalendar.DisplayDate;

                        distance = new Dictionary<(string supRegion, string cusRegion), int>(4)
                        {
                            {("MB", "MB"), int.Parse(NorthNorth.Text)},
                            {("MN", "MN"), int.Parse(SouthSouth.Text)},
                            {("LD", "MB"), int.Parse(MidNorth.Text)},
                            {("LD", "MN"), int.Parse(MidSouth.Text)}
                        };
                    }));

                int maxDistance = distance.Values.Max();

                dateTo = dateTo > dateFrom
                    ? dateTo
                    : dateFrom;

                var products = new Dictionary<string, Product>();
                var suppliers = new Dictionary<string, Supplier>();

                // var dicFc = new Dictionary<(DateTime DateFc, string SupplierCode, string ProductCode), (SupplierForecast Supply, bool)>();

                // Date Forecast - Region - ProductCode - Supply & Valid
                var dicFc = new Dictionary<DateTime, Dictionary<string, Dictionary<string, Dictionary<SupplierForecast, bool>>>>();

                var customers = new Dictionary<string, Customer>();

                //// var dicPo = new Dictionary<(DateTime DatePo, string CusKeyCode, string ProductCode), (CustomerOrder Order, bool)>();

                // Date Order - Region - ProductCode - Order & Valid
                var dicPo =
                    new Dictionary<DateTime, Dictionary<string, Dictionary<string, Dictionary<CustomerOrder, bool>>>>();

                var dicMoq = new Dictionary<string, double>
                {
                    {"K01901", 0.3}, // Chanh có hạt
                    {"K02201", 0.3}, // Chanh không hạt
                    {"C07101", 0.1}, // Ớt ngọt ( chuông ) đỏ
                    {"C07201", 0.1}, // Ớt ngọt ( chuông ) vàng
                    {"C07301", 0.1}, // Ớt ngọt ( chuông ) xanh
                    {"B00201", 0.3}, // Dọc mùng ( bạc hà )
                    {"C01801", 0.3}, // Cà chua cherry đỏ
                    {"C04401", 0.3} // Đậu bắp xanh
                };

                WriteToRichTextBoxOutput(
                    "Bắt đầu đọc Data",
                    1);

                var readTasks = new[]
                {
                    // Products
                    new Task(delegate { products = ReadProducts(); }),

                    // Suppliers
                    new Task(delegate { suppliers = ReadSuppliers(); }),

                    // Customers
                    new Task(delegate { customers = ReadCustomers(); })
                };

                // Here we go.
                Parallel.ForEach(
                    readTasks,
                    new ParallelOptions
                    {
                        MaxDegreeOfParallelism = Environment.ProcessorCount
                    },
                    task => { task.Start(); });

                // Gonna wait for all reading tasks to finish.
                await Task.WhenAll(readTasks).ConfigureAwait(true);
                WriteToRichTextBoxOutput(
                    "Đọc xong Data - Phần 1.",
                    2);

                readTasks = new[]
                {
                    // Forecasts
                    new Task(
                        delegate
                        {
                            // Safeguard
                            if (!File.Exists($@"{_applicationPath}\Database\Forecasts.xlsb"))
                            {
                                WriteToRichTextBoxOutput("Không có Database Forecast.");
                                return;
                            }

                            using (var workbook = new Workbook(
                                $@"{_applicationPath}\Database\Forecasts.xlsb",
                                new LoadOptions
                                {
                                    MemorySetting = MemorySetting.MemoryPreference
                                }))
                            {
                                Worksheet worksheet = workbook.Worksheets[0];
                                using (DataTable table = worksheet.Cells.ExportDataTable(
                                    0,
                                    0,
                                    worksheet.Cells.MaxDataRow + 1,
                                    worksheet.Cells.MaxDataColumn + 1,
                                    _globalExportTableOptionsOpts))
                                {
                                    var colFirst = 0;
                                    var colLast = 0;

                                    for (var colIndex = 0;
                                        colIndex < table.Columns.Count;
                                        colIndex++)
                                        using (DataColumn column = table.Columns[colIndex])
                                        {
                                            DateTime? dateFc = _ulti.StringToDate(
                                                _ulti.GetString(column.ColumnName));
                                            {
                                                if (dateFc == null) continue;
                                            }

                                            if (dateFc == dateFrom.AddDays(-maxDistance)) colFirst = colIndex;

                                            if (dateFc != dateTo.AddDays(maxDistance)) continue;

                                            colLast = colIndex;
                                            break;
                                        }

                                    foreach (DataRow row in table.Select())
                                    {
                                        string productCode = _ulti.ObjectToString(row["ProductCode"]);
                                        string supplierCode = _ulti.ObjectToString(row["SupplierCode"]);

                                        for (int colIndex = colFirst;
                                            colIndex <= colLast;
                                            colIndex++)
                                            using (DataColumn column = table.Columns[colIndex])
                                            {
                                                // First check point. Is it a valid date?
                                                DateTime? dateFc =
                                                    _ulti.StringToDate(column.ColumnName);

                                                // FiteMoi specific Validation for date.
                                                // if (dateFc == null || dateFc > dateTo || dateFc < dateFrom) continue;

                                                // Second check point. Is it a valid forecast value?
                                                double value = _ulti.ObjectToDouble(row[colIndex]);
                                                if (value <= 0) continue;

                                                // ReSharper disable once PossibleInvalidOperationException
                                                if (!dicFc.TryGetValue((DateTime) dateFc, out Dictionary<string, Dictionary<string, Dictionary<SupplierForecast, bool>>> fcRegion))
                                                {
                                                    fcRegion = new Dictionary<string, Dictionary<string, Dictionary<SupplierForecast, bool>>>();
                                                    dicFc.Add(
                                                        (DateTime) dateFc,
                                                        fcRegion);
                                                }

                                                if (!fcRegion.TryGetValue(suppliers[supplierCode].SupplierRegion, out Dictionary<string, Dictionary<SupplierForecast, bool>> fcProducts))
                                                {
                                                    fcProducts = new Dictionary<string, Dictionary<SupplierForecast, bool>>();
                                                    fcRegion.Add(
                                                        suppliers[supplierCode]
                                                            .SupplierRegion,
                                                        fcProducts);
                                                }

                                                if (!fcProducts.TryGetValue(productCode, out Dictionary<SupplierForecast, bool> fcSupplies))
                                                {
                                                    fcSupplies = new Dictionary<SupplierForecast, bool>();
                                                    fcProducts.Add(
                                                        productCode,
                                                        fcSupplies);
                                                }

                                                fcSupplies.Add(
                                                    new SupplierForecast
                                                    {
                                                        QualityControlPass = true,
                                                        SupplierCode = supplierCode,
                                                        FullOrder = _ulti.ObjectToInt(row["FullOrder"]) == 1,
                                                        CrossRegion = _ulti.ObjectToInt(row["CrossRegion"]) == 1,
                                                        LabelVinEco = _ulti.ObjectToInt(row["Label"]) == 1,
                                                        Level = (byte) _ulti.ObjectToInt(row["Level"])
                                                    },
                                                    false);
                                            }
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
                                string path = $@"{_applicationPath}\Database\Orders.xlsb";
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
                                        worksheet.Cells.MaxDataRow + 1,
                                        worksheet.Cells.MaxDataColumn + 1,
                                        _globalExportTableOptionsOpts))
                                    {
                                        var colFirst = 0;
                                        var colLast = 0;

                                        for (var colIndex = 0;
                                            colIndex < table.Columns.Count;
                                            colIndex++)
                                            using (DataColumn column = table.Columns[colIndex])
                                            {
                                                DateTime? dateFc = _ulti.StringToDate(
                                                    _ulti.GetString(column.ColumnName));
                                                if (dateFc == null) continue;

                                                if (dateFc == dateFrom.AddDays(-maxDistance)) colFirst = colIndex;

                                                if (dateFc != dateTo.AddDays(maxDistance)) continue;

                                                // Once encounter dateTo, break, and record its location.
                                                // Further optimization, coz it IS my database.
                                                colLast = colIndex;
                                                break;
                                            }

                                        foreach (DataRow row in table.Select())
                                        {
                                            string productCode =
                                                _ulti.ObjectToString(row["ProductCode"]);
                                            string cusKeyCode =
                                                _ulti.ObjectToString(row["CustomerKeyCode"]);

                                            for (int colIndex = colFirst;
                                                colIndex <= colLast;
                                                colIndex++)
                                                using (DataColumn column = table.Columns[colIndex])
                                                {
                                                    // First check point. Is it a valid date?
                                                    // ReSharper disable once PossibleInvalidOperationException
                                                    // Because I'm confident about that.
                                                    // ... it's my fucking database.
                                                    DateTime? datePo = _ulti.StringToDate(
                                                        _ulti.GetString(column.ColumnName));

                                                    // if (datePo == null || datePo > dateTo || datePo < dateFrom) continue;

                                                    // Second check point. Is it a valid forecast value?
                                                    double value = _ulti.ObjectToDouble(row[colIndex]);
                                                    if (value <= 0) continue;

                                                    // ReSharper disable once PossibleInvalidOperationException
                                                    if (!dicPo.TryGetValue((DateTime) datePo, out Dictionary<string, Dictionary<string, Dictionary<CustomerOrder, bool>>> poRegion))
                                                    {
                                                        poRegion = new Dictionary<string, Dictionary<string, Dictionary<CustomerOrder, bool>>>();
                                                        dicPo.Add(
                                                            (DateTime) datePo,
                                                            poRegion);
                                                    }

                                                    if (!poRegion.TryGetValue(customers[cusKeyCode].CustomerBigRegion, out Dictionary<string, Dictionary<CustomerOrder, bool>> poProducts))
                                                    {
                                                        poProducts = new Dictionary<string, Dictionary<CustomerOrder, bool>>();
                                                        poRegion.Add(
                                                            customers[cusKeyCode]
                                                                .CustomerBigRegion,
                                                            poProducts);
                                                    }

                                                    if (!poProducts.TryGetValue(productCode, out Dictionary<CustomerOrder, bool> poOrders))
                                                    {
                                                        poOrders = new Dictionary<CustomerOrder, bool>();
                                                        poProducts.Add(
                                                            productCode,
                                                            poOrders);
                                                    }

                                                    poOrders.Add(
                                                        new CustomerOrder
                                                        {
                                                            CustomerKeyCode =
                                                                cusKeyCode,
                                                            QuantityOrder = value
                                                        },
                                                        false);

                                                    // dicPo.Add(
                                                    // // ReSharper disable once PossibleInvalidOperationException
                                                    // ((DateTime)datePo, cusKeyCode, productCode
                                                    // ),
                                                    // (new CustomerOrder
                                                    // {
                                                    // CustomerKeyCode =
                                                    // cusKeyCode,
                                                    // QuantityOrder =
                                                    // poValue
                                                    // }, false));
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
                Parallel.ForEach(
                    readTasks,
                    new ParallelOptions
                    {
                        MaxDegreeOfParallelism = Environment.ProcessorCount
                    },
                    task => { task.Start(); });

                // Gonna wait for all reading tasks to finish.
                await Task.WhenAll(readTasks).ConfigureAwait(true);
                WriteToRichTextBoxOutput(
                    "Đọc xong Data - Phần 2.",
                    2);

                List<DateTime> listDatePo = dicPo.Keys.ToList();

                WriteToRichTextBoxOutput("Here goes nothing.");

                var coordResult = new Dictionary<string, Dictionary<(DateTime DatePo, CustomerOrder Order), (DateTime DateFc, SupplierForecast Supply)>>();

                foreach (string priority in new[]
                {
                    "B2B",

                    // "VM+ VinEco Priority",
                    "VM+ VinEco",

                    // "VM Priority",
                    // "VM+",
                    // "VM",
                    string.Empty
                })
                    Coord(priority);

                void Coord(
                    string priority = "")
                {
                    try
                    {
                        var localCustomers = new ConcurrentDictionary<string, Customer>(customers);
                        var localSuppliers = new ConcurrentDictionary<string, Supplier>(suppliers);
                        var localWatch = new Stopwatch();

                        foreach (DateTime datePo in listDatePo)
                        {
                            localWatch.Restart();

                            foreach (string productCode in products.Keys)
                            {
                                Dictionary<CustomerOrder, bool> GetOrderDictionary(DateTime date, string region)
                                {
                                    if (dicPo.TryGetValue(date,
                                            out Dictionary<string, Dictionary<string, Dictionary<CustomerOrder, bool>>>
                                                orderRegions) &&
                                        orderRegions.TryGetValue(region,
                                            out Dictionary<string, Dictionary<CustomerOrder, bool>>
                                                orderRegionProducts) &&
                                        orderRegionProducts.TryGetValue(productCode,
                                            out Dictionary<CustomerOrder, bool> result))
                                    {
                                        return result.Where(
                                                po => priority == string.Empty ||
                                                      localCustomers[po.Key.CustomerKeyCode]
                                                          .CustomerType ==
                                                      priority)
                                            .ToDictionary(
                                                po => po.Key,
                                                po => false);
                                    }

                                    return null;
                                }
                                
                                Dictionary<CustomerOrder, bool> orderNorth = GetOrderDictionary(datePo, "MB");

                                Dictionary<CustomerOrder, bool> orderSouth = GetOrderDictionary(
                                    datePo.AddDays(-distance[("LD", "MB")] + distance[("LD", "MN")]),
                                    "MN");

                                double SumOrder(Dictionary<CustomerOrder, bool> source)
                                {
                                    return source?.AsParallel()
                                               .Sum(po => po.Key.QuantityOrder) ??
                                           0;
                                }

                                double sumPoNorth = SumOrder(orderNorth);
                                double sumPoSouth = SumOrder(orderSouth);

                                Dictionary<SupplierForecast, bool> GetForecastDictionary(DateTime date, string region)
                                {
                                    if (dicFc.TryGetValue(date, out Dictionary<string, Dictionary<string, Dictionary<SupplierForecast, bool>>> forecastRegions) && forecastRegions.TryGetValue(region, out Dictionary<string, Dictionary<SupplierForecast, bool>> forecastRegionProducts) && forecastRegionProducts.TryGetValue(productCode, out Dictionary<SupplierForecast, bool> result))
                                        return result.OrderBy(fc => fc.Key.QuantityForecast)
                                            .ToDictionary(
                                                fc => fc.Key,
                                                fc => true);

                                    return null;
                                }

                                Dictionary<SupplierForecast, bool> forecastNorth = GetForecastDictionary(
                                    datePo.AddDays(distance[("MB", "MB")]),
                                    "MB");

                                Dictionary<SupplierForecast, bool> forecastMid = GetForecastDictionary(
                                    datePo.AddDays(
                                        -Math.Max(
                                            distance[("LD", "MB")],
                                            distance[("LD", "MN")])),
                                    "LD");

                                Dictionary<SupplierForecast, bool> forecastSouth = GetForecastDictionary(
                                    datePo.AddDays(-distance[("MN", "MN")]),
                                    "MN");

                                double SumForecast(Dictionary<SupplierForecast, bool> source)
                                {
                                    return source?.AsParallel().Sum(fc => fc.Key.QuantityForecast) ?? 0;
                                }

                                double sumFcNorth = SumForecast(forecastNorth);
                                double sumFcMid = SumForecast(forecastMid);
                                double sumFcSouth = SumForecast(forecastSouth);

                                // Todo - Implement Rate
                                // Working on this.
                                CalculateRate();

                                void CalculateRate()
                                {
                                    try
                                    {
                                        double northMissing = sumPoNorth - sumFcNorth;
                                        double southMissing = sumPoSouth - sumFcSouth;

                                        double rateNorth = (sumFcNorth + sumFcMid * (northMissing / (northMissing + southMissing))) / sumPoNorth;
                                        double rateSouth = (sumFcSouth + sumFcMid * (southMissing / (northMissing + southMissing))) / sumPoSouth;

                                        void PairSupplyOrder(CustomerOrder customerOrder,
                                            IDictionary<CustomerOrder, bool> orders,
                                            IDictionary<SupplierForecast, bool> forecasts, double rate)
                                        {
                                            try
                                            {
                                                // Validation.
                                                // Why the heck is this empty in the first place?
                                                if (!forecasts.Any()) return;

                                                SupplierForecast supply = forecasts.Aggregate((current, next) => current.Key.QuantityForecast > next.Key.QuantityForecast ? current : next).Key;
                                                forecasts.Remove(supply);

                                                double deliQuantity = Math.Min(
                                                    customerOrder.QuantityOrder * rate,
                                                    supply.QuantityForecast);

                                                supply.QuantityForecast -= deliQuantity;

                                                // ReSharper disable once SuggestVarOrType_Elsewhere
                                                if (!coordResult.TryGetValue(productCode, out var dicCoord))
                                                {
                                                    dicCoord = new Dictionary<(DateTime DatePo, CustomerOrder Order), (DateTime DateFc, SupplierForecast Supply)>();
                                                    coordResult.Add(productCode, dicCoord);
                                                }

                                                dicCoord.Add(
                                                    (datePo, customerOrder),
                                                    (datePo.AddDays(-1), supply));

                                                forecasts.Add(supply, false);

                                                orders.Remove(customerOrder);
                                            }
                                            catch (Exception ex)
                                            {
                                                WriteToRichTextBoxOutput(ex.Message);
                                                throw;
                                            }
                                        }

                                        if (orderNorth != null && sumFcNorth + sumFcMid >= 0)
                                        {
                                            foreach (CustomerOrder customerOrder in orderNorth.Keys.OrderByDescending(po => po.QuantityOrder).ToList())
                                            {
                                                if (forecastNorth != null)
                                                {
                                                    PairSupplyOrder(customerOrder, orderNorth, forecastNorth, rateNorth);

                                                    continue;
                                                }

                                                if (forecastMid == null) continue;

                                                {
                                                    PairSupplyOrder(customerOrder, orderNorth, forecastMid, rateNorth);
                                                }
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        WriteToRichTextBoxOutput(ex.Message);
                                        throw;
                                    }
                                }

                                // Todo - Select Supplier for each Order.
                            }
                        }

                        WriteToRichTextBoxOutput(
                            $"{(priority == string.Empty ? "Còn lại" : priority)}: {Math.Round(localWatch.Elapsed.TotalSeconds, 2).ToString(CultureInfo.InvariantCulture)}s!");

                        localWatch.Stop();

                        // Export stuff to Excel files.
                        // Otherwise, what did we do all of these for? lol.
                    }
                    catch (Exception ex)
                    {
                        WriteToRichTextBoxOutput(ex.Message);
                        throw;
                    }
                }
                
                ExportResult(coordResult, products);

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

        /// <summary>
        ///     Exporting CoordResult to Excel Files.
        /// </summary>
        private void ExportResult(
            Dictionary<string,
                Dictionary<(DateTime DatePo, CustomerOrder Order),
                    (DateTime DateFc, SupplierForecast Supply)>> coordResult,
            Dictionary<string, Product> products,
            Dictionary<string, Customer> customers,
            Dictionary<string, Supplier> suppliers)
        {
            try
            {
                // Todo - Do this. Obviously.
                string YesNoFromString(bool expression)
                {
                    return expression ? "Yes" : "No";
                }

                string StringFromNullableString(string source)
                {
                    return source ?? "Any";
                }

                DataTable tableMastahCompact = ExportMastahCompact();

                DataTable ExportMastahCompact()
                {
                    try
                    {
                        using (var table = new DataTable {TableName = "Mastah Compact"})
                        {
                            foreach ((string columnName, Type columnType, object columnDefaultValue) details in new[]
                            {
                                ("Mã 6 ký tự", typeof(string), null),
                                ("Tên sản phẩm", typeof(string), null),
                                ("ProductOrientation", typeof(string), null),
                                ("ProductClimate", typeof(string), null),
                                ("ProductionGroup", typeof(string), null),
                                ("Nhóm sản phẩm", typeof(string), null),
                                ("Ghi chú", typeof(string), null),
                                ("Loại cửa hàng", typeof(string), null),
                                ("P&L", typeof(string), null),
                                ("Ngày tiêu thụ", typeof(DateTime), null),
                                ("Tỉnh tiêu thụ", typeof(string), null),
                                ("Vùng tiêu thụ", typeof(string), null),
                                ("Vùng SX yêu cầu", typeof(string), null),
                                ("Nguồn yêu cầu", typeof(string), null),
                                ("Nhu cầu", typeof(double), _ulti.DoubleToObject(0)),
                                ("Đáp ứng", typeof(double), _ulti.DoubleToObject(0)),
                                ("Nguồn", typeof(string), null),
                                ("Vùng sản xuất", typeof(string), null),
                                ("Mã NCC", typeof(string), null),
                                ("Tên NCC", typeof(string), null),
                                ("Ngày sơ chế", typeof(DateTime), null),
                                ("NoSup", typeof(double), _ulti.DoubleToObject(0)),
                                ("KPI", typeof(double), _ulti.DoubleToObject(0)),
                                ("Label", typeof(string), null),
                                ("CodeSFG", typeof(string), null),
                                ("IsNoSup", typeof(bool), _ulti.BoolToObject(false)),
                            })
                            {
                                table.Columns.Add(details.columnName, details.columnType).DefaultValue = details.columnDefaultValue;
                            }

                            // Dictionaries for row.
                            // Just in case, dupplicated row happens.
                            // In which case, happens a fucking lot for Mastah Compact
                            var dicRow = new Dictionary<string, int>();

                            foreach (string productCode in coordResult.Keys)
                            {
                                foreach (KeyValuePair<(DateTime DatePo, CustomerOrder Order), (DateTime DateFc, SupplierForecast Supply)> pair in coordResult[productCode)
                                {
                                    Product product = products[productCode];
                                    Customer customer = customers[pair.Key.Order.CustomerKeyCode];
                                    Supplier supplier = suppliers[pair.Value.Supply.SupplierCode];
                                    
                                    // Building 'unique' rowKey to identify rows.
                                    string rowKey = $"{_ulti.DateToString(pair.Key.DatePo, "yyyyMMdd")}-{customer.CustomerKeyCode}-{_ulti.DateToString(pair.Value.DateFc, "yyyyMMdd")}-{supplier.SupplierCode}";

                                    // Initializing
                                    DataRow dr;
                                       
                                    // ... And check if row exists yet
                                    if (!dicRow.TryGetValue(rowKey, out int rowIndex))
                                    {
                                        // If not.
                                        dr = table.NewRow();
                                        dicRow.Add(rowKey, table.Rows.Count);
                                        table.Rows.Add(dr);
                                        dr = table.Rows[table.Rows.Count - 1];
                                    }
                                    else
                                    {
                                        // If exists.
                                        dr = table.Rows[rowIndex];

                                        dr["Mã 6 ký tự"] = productCode;
                                        dr["Tên sản phẩm"] = product.ProductName;
                                        dr["Nhóm sản phẩm"] = product.ProductClassification;
                                        dr["ProductOrientation"] = product.ProductOrientation;
                                        dr["ProductClimate"] = product.ProductClimate;
                                        dr["ProductionGroup"] = product.ProductionGroup;
                                        dr["Ghi chú"] =
                                            product.ProductNote.Contains(customer.CustomerBigRegion == "Miền Nam"
                                                ? "South"
                                                : "North")
                                                ? "Ok"
                                                : "Out of List";
                                        dr["Loại cửa hàng"] = customer.CustomerType;
                                        dr["P&L"] = customer.Company;
                                        //dr["Ngày tiêu thụ"] = (int)(DatePO.Date - _dateBase).TotalDays + 2;
                                        dr["Ngày tiêu thụ"] = _ulti.DateToObject(pair.Key.DatePo, "yyyyMMdd");
                                        dr["Vùng tiêu thụ"] = customer.CustomerBigRegion;
                                        // Todo - Add YesNoSubRegion here.
                                        // dr["Tỉnh tiêu thụ"] = YesNoSubRegion ? customer.CustomerRegion : null;
                                        dr["Vùng SX yêu cầu"] = StringFromNullableString(pair.Key.Order.DesiredRegion);
                                        dr["Nguồn yêu cầu"] = StringFromNullableString(pair.Key.Order.DesiredSource);

                                        dr["Nhu cầu"] = _ulti.DoubleToObject((double) dr["Nhu cầu"] + pair.Key.Order.QuantityOrder);
                                        dr["Đáp ứng"] = _ulti.DoubleToObject((double) dr["Đáp ứng"] + pair.Value.Supply.QuantityForecast);
                                    }

                                    if (pair.Key.Order.QuantityOrder > 0)
                                    {
                                        dr["Nguồn"] = supplier.SupplierType;
                                        dr["Vùng sản xuất"] = supplier.SupplierRegion;
                                        dr["Mã NCC"] = supplier.SupplierCode;
                                        dr["Tên NCC"] = supplier.SupplierName;
                                        dr["Ngày sơ chế"] = _ulti.DateToObject(pair.Value.DateFc, "yyyyMMdd");
                                        dr["Label"] = YesNoFromString(pair.Value.Supply.LabelVinEco);
                                        dr["CodeSFG"] = $"{productCode}1{_ulti.IntToObject((supplier.SupplierRegion == "Lâm Đồng" ? 0 : 2) + (pair.Value.Supply.LabelVinEco ? 1 : 0))}";
                                    }
                                    else
                                    {
                                        dr["Nguồn"] = "Không đáp ứng";
                                    }
                                }
                            }
                            
                            foreach (DataRow dr in table.Select())
                            {
                                dr["NoSup"] = _ulti.DoubleToObject(_ulti.ZeroIfNegative(dr["Nhu cầu"], dr["Đáp ứng"]));
                                if ((double) dr["NoSup"] > 1)
                                {
                                    dr["IsNoSup"] = _ulti.BoolToObject(true);
                                }
                            }

                            return table;
                        }
                    }
                    catch (Exception ex)
                    {
                        WriteToRichTextBoxOutput(ex.Message);
                        throw;
                    }
                }

                foreach (DataRow dr in table.Rows)
                {
                    dr["NoSup"] = Math.Max((double) dr["Nhu cầu"] - (double) dr["Đáp ứng"], 0);
                    if ((double) dr["NoSup"] > 1) dr["IsNoSup"] = true;
                }

                IOrderedEnumerable<ForecastDate> _FC = db.GetCollection<ForecastDate>("Forecast")
                    .Find(x =>
                        x.DateForecast >= DateFrom.Date &&
                        x.DateForecast <= DateTo.Date)
                    .ToList()
                    .OrderByDescending(x => x.DateForecast);

                foreach (ForecastDate _ForecastDate in _FC)
                foreach (ProductForecast _ProductForecast in _ForecastDate.ListProductForecast)
                foreach (SupplierForecast _SupplierForecast in _ProductForecast.ListSupplierForecast.Where(
                    x =>
                        x.QualityControlPass && x.QuantityForecastPlanned > 0))
                {
                    DataRow dr = table.NewRow();

                    Product _Product = coreStructure.dicProduct[_ProductForecast.ProductId];
                    Supplier _Supplier = coreStructure.dicSupplier[_SupplierForecast.SupplierId];

                    if (FruitOnly)
                        if (_Product.ProductCode != "K" && _Product.ProductCode != "D01401")
                            continue;

                    dr["Mã 6 ký tự"] = _Product.ProductCode;
                    dr["Tên sản phẩm"] = _Product.ProductName;
                    dr["Nhóm sản phẩm"] = _Product.ProductClassification;
                    dr["ProductOrientation"] = _Product.ProductOrientation;
                    dr["ProductClimate"] = _Product.ProductClimate;
                    dr["ProductionGroup"] = _Product.ProductionGroup;
                    dr["Ghi chú"] = _Product.ProductNote.Count != 0 ? "Ok" : "Out of List";

                    dr["Nguồn"] = _Supplier.SupplierType;
                    dr["Vùng sản xuất"] = _Supplier.SupplierRegion;
                    dr["Mã NCC"] = _Supplier.SupplierCode;
                    dr["Tên NCC"] = _Supplier.SupplierName;

                    //dr["Ngày sơ chế"] = (int)(_ForecastDate.DateForecast.Date - _dateBase).TotalDays + 2;
                    dr["Ngày sơ chế"] = _ForecastDate.DateForecast.Date;
                    dr["KPI"] = _SupplierForecast.QuantityForecastPlanned;

                    table.Rows.Add(dr);
                }

                #endregion
            }
            catch (Exception ex)
            {
                WriteToRichTextBoxOutput(ex.Message);
                throw;
            }
        }

        /// <summary>
        ///     Reading Customers from database.
        /// </summary>
        /// <returns>
        ///     The <see cref="Task" />.
        /// </returns>
        private Dictionary<string, Customer> ReadCustomers()
        {
            try
            {
                var customers = new Dictionary<string, Customer>();
                if (!File.Exists($@"{_applicationPath}\Database\Customers.xlsb")) return null;

                using (var workbook = new Workbook(
                    $@"{_applicationPath}\Database\Customers.xlsb",
                    new LoadOptions
                    {
                        MemorySetting = MemorySetting.MemoryPreference
                    }))
                {
                    Worksheet worksheet = workbook.Worksheets[0];
                    using (DataTable table = worksheet.Cells.ExportDataTable(
                        0,
                        0,
                        worksheet.Cells.MaxDataRow + 1,
                        worksheet.Cells.MaxDataColumn + 1,
                        _globalExportTableOptionsOpts))
                    {
                        foreach (DataRow row in table.Select())
                            customers.Add(
                                _ulti.ObjectToString(row["Key"]),
                                new Customer
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

                return customers;
            }
            catch (Exception ex)
            {
                WriteToRichTextBoxOutput(ex.Message);
                throw;
            }
        }

        /// <summary>
        ///     Reading Products from database.
        /// </summary>
        /// <returns>
        ///     The <see cref="Task" />.
        /// </returns>
        private Dictionary<string, Product> ReadProducts()
        {
            try
            {
                // Products
                var products = new Dictionary<string, Product>();
                if (!File.Exists($@"{_applicationPath}\Database\Products.xlsb")) return null;

                using (var workbook = new Workbook(
                    $@"{_applicationPath}\Database\Products.xlsb",
                    new LoadOptions
                    {
                        MemorySetting = MemorySetting.MemoryPreference
                    }))
                {
                    Worksheet worksheet = workbook.Worksheets[0];
                    using (DataTable table = worksheet.Cells.ExportDataTable(
                        0,
                        0,
                        worksheet.Cells.MaxDataRow + 1,
                        worksheet.Cells.MaxDataColumn + 1,
                        _globalExportTableOptionsOpts))
                    {
                        foreach (DataRow row in table.Select())
                            products.Add(
                                _ulti.ObjectToString(row["ProductCode"]),
                                new Product
                                {
                                    ProductCode = _ulti.ObjectToString(row["ProductCode"]),
                                    ProductName = _ulti.ObjectToString(row["ProductName"])
                                });
                    }
                }

                return products;
            }
            catch (Exception ex)
            {
                WriteToRichTextBoxOutput(ex.Message);
                throw;
            }
        }

        /// <summary>
        ///     Reading Suppliers from database.
        /// </summary>
        /// <returns>
        ///     The <see cref="Task" />.
        /// </returns>
        private Dictionary<string, Supplier> ReadSuppliers()
        {
            try
            {
                var suppliers = new Dictionary<string, Supplier>();
                if (!File.Exists($@"{_applicationPath}\Database\Suppliers.xlsb")) return null;

                using (var workbook = new Workbook(
                    $@"{_applicationPath}\Database\Suppliers.xlsb",
                    new LoadOptions
                    {
                        MemorySetting = MemorySetting.MemoryPreference
                    }))
                {
                    Worksheet worksheet = workbook.Worksheets[0];
                    using (DataTable table = worksheet.Cells.ExportDataTable(
                        0,
                        0,
                        worksheet.Cells.MaxDataRow + 1,
                        worksheet.Cells.MaxDataColumn + 1,
                        _globalExportTableOptionsOpts))
                    {
                        foreach (DataRow row in table.Select())
                            suppliers.Add(
                                _ulti.ObjectToString(row["SupplierCode"]),
                                new Supplier
                                {
                                    SupplierRegion = _ulti.ObjectToString(row["SupplierRegion"]),
                                    SupplierType = _ulti.ObjectToString(row["SupplierType"]),
                                    SupplierCode = _ulti.ObjectToString(row["SupplierCode"]),
                                    SupplierName = _ulti.ObjectToString(row["SupplierName"])
                                });
                    }
                }

                return suppliers;
            }
            catch (Exception ex)
            {
                WriteToRichTextBoxOutput(ex.Message);
                throw;
            }
        }
    }
}