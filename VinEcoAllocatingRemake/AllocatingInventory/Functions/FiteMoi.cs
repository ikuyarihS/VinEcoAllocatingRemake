// --------------------------------------------------------------------------------------------------------------------
// <copyright file="FiteMoi.cs" company="VinEco">
//   Shirayuki 2018.
// </copyright>
// <summary>
//   The allocating inventory.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

using DocumentFormat.OpenXml.Wordprocessing;

namespace VinEcoAllocatingRemake.AllocatingInventory
{
    #region

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

    #endregion

    /// <summary>
    ///     The allocating inventory.
    /// </summary>
    // ReSharper disable once StyleCop.SA1404
    [SuppressMessage("ReSharper", "ArrangeThisQualifier")]
    public partial class AllocatingInventory
    {
        /// <summary>
        ///     Fite moi!.
        /// </summary>
        /// <param name="sender"> The sender. </param>
        /// <param name="e"> The e. </param>
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
                            { ("MB", "MB"), 1 },
                            { ("MN", "MN"), 0 },
                            { ("LD", "MB"), 3 },
                            { ("LD", "MN"), 0 }
                        };

                // ReSharper disable once AsyncConverter.ConfigureAwaitHighlighting
                await this.Dispatcher.BeginInvoke(
                    (Action)(() =>
                        {
                            dateFrom = this.DateFromCalendar.SelectedDate ?? DateTime.Today;
                            dateTo = this.DateToCalendar.SelectedDate ?? DateTime.Today;

                            distance = new Dictionary<(string supRegion, string cusRegion), int>(4)
                                           {
                                               { ("MB", "MB"), int.Parse(this.NorthNorth.Text) },
                                               { ("MN", "MN"), int.Parse(this.SouthSouth.Text) },
                                               { ("LD", "MB"), int.Parse(this.MidNorth.Text) },
                                               { ("LD", "MN"), int.Parse(this.MidSouth.Text) }
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
                                     { "K01901", 0.3 }, // Chanh có hạt
                                     { "K02201", 0.3 }, // Chanh không hạt
                                     { "C07101", 0.1 }, // Ớt ngọt ( chuông ) đỏ
                                     { "C07201", 0.1 }, // Ớt ngọt ( chuông ) vàng
                                     { "C07301", 0.1 }, // Ớt ngọt ( chuông ) xanh
                                     { "B00201", 0.3 }, // Dọc mùng ( bạc hà )
                                     { "C01801", 0.3 }, // Cà chua cherry đỏ
                                     { "C04401", 0.3 }, // Đậu bắp xanh
                                 };

                this.WriteToRichTextBoxOutput(
                    "Bắt đầu đọc Data",
                    1);

                var readTasks = new[]
                                    {
                                        // Products
                                        new Task(delegate { products = this.ReadProducts(); }),

                                        // Suppliers
                                        new Task(delegate { suppliers = this.ReadSuppliers(); }),

                                        // Customers
                                        new Task(delegate { customers = this.ReadCustomers(); })
                                    };

                // Here we go.
                Parallel.ForEach(
                    readTasks,
                    new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount },
                    task => { task.Start(); });

                // Gonna wait for all reading tasks to finish.
                await Task.WhenAll(readTasks).ConfigureAwait(true);
                this.WriteToRichTextBoxOutput(
                    "Đọc xong Data - Phần 1.",
                    2);

                readTasks = new[]
                                {
                                    // Forecasts
                                    new Task(
                                        delegate
                                            {
                                                // Safeguard
                                                if (!File.Exists($@"{this.applicationPath}\Database\Forecasts.xlsb"))
                                                {
                                                    this.WriteToRichTextBoxOutput("Không có Database Forecast.");
                                                    return;
                                                }

                                                using (var workbook = new Workbook(
                                                    $@"{this.applicationPath}\Database\Forecasts.xlsb",
                                                    new LoadOptions { MemorySetting = MemorySetting.MemoryPreference }))
                                                {
                                                    Worksheet worksheet = workbook.Worksheets[0];
                                                    using (DataTable table = worksheet.Cells.ExportDataTable(
                                                        0,
                                                        0,
                                                        worksheet.Cells.MaxDataRow + 1,
                                                        worksheet.Cells.MaxDataColumn + 1,
                                                        this.globalExportTableOptionsOpts))
                                                    {
                                                        var colFirst = 0;
                                                        var colLast = 0;

                                                        for (var colIndex = 0;
                                                             colIndex < table.Columns.Count;
                                                             colIndex++)
                                                        {
                                                            using (DataColumn column = table.Columns[colIndex])
                                                            {
                                                                DateTime? dateFc = this.ulti.StringToDate(this.ulti.GetString(column.ColumnName));

                                                                if (dateFc == null)
                                                                {
                                                                    continue;
                                                                }

                                                                if (dateFc == dateFrom.AddDays(-maxDistance))
                                                                {
                                                                    colFirst = colIndex;
                                                                }

                                                                if (dateFc != dateTo.AddDays(maxDistance))
                                                                {
                                                                    continue;
                                                                }

                                                                colLast = colIndex;
                                                                break;
                                                            }
                                                        }

                                                        foreach (DataRow row in table.Select())
                                                        {
                                                            string productCode = this.ulti.ObjectToString(row["ProductCode"]);
                                                            string supplierCode = this.ulti.ObjectToString(row["SupplierCode"]);

                                                            for (int colIndex = colFirst;
                                                                 colIndex <= colLast;
                                                                 colIndex++)
                                                            {
                                                                using (DataColumn column = table.Columns[colIndex])
                                                                {
                                                                    // First check point. Is it a valid date?
                                                                    DateTime? dateFc = this.ulti.StringToDate(column.ColumnName);

                                                                    // FiteMoi specific Validation for date.
                                                                    ////if (dateFc == null ||
                                                                    ////    dateFc > dateTo.AddDays(maxDistance) ||
                                                                    ////    dateFc < dateFrom.AddDays(-maxDistance))
                                                                    ////{
                                                                    ////    continue;
                                                                    ////}

                                                                    // Second check point. Is it a valid forecast value?
                                                                    double value = this.ulti.ObjectToDouble(row[colIndex]);
                                                                    if (value <= 0)
                                                                    {
                                                                        continue;
                                                                    }

                                                                    // ReSharper disable once PossibleInvalidOperationException
                                                                    if (!dicFc.TryGetValue((DateTime)dateFc, out Dictionary<string, Dictionary<string, Dictionary<SupplierForecast, bool>>> fcRegion))
                                                                    {
                                                                        fcRegion = new Dictionary<string, Dictionary<string, Dictionary<SupplierForecast, bool>>>();
                                                                        dicFc.Add(
                                                                            (DateTime)dateFc,
                                                                            fcRegion);
                                                                    }

                                                                    if (!fcRegion.TryGetValue(suppliers[supplierCode].SupplierRegion, out Dictionary<string, Dictionary<SupplierForecast, bool>> fcProducts))
                                                                    {
                                                                        fcProducts = new Dictionary<string, Dictionary<SupplierForecast, bool>>();
                                                                        fcRegion.Add(
                                                                            suppliers[supplierCode].SupplierRegion,
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
                                                                                FullOrder = this.ulti.ObjectToInt(row["FullOrder"]) == 1,
                                                                                CrossRegion = this.ulti.ObjectToInt(row["CrossRegion"]) == 1,
                                                                                LabelVinEco = this.ulti.ObjectToInt(row["Label"]) == 1,
                                                                                Level = (byte)this.ulti.ObjectToInt(row["Level"]),
                                                                                QuantityForecast = value
                                                                            },
                                                                        false);
                                                                }
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
                                                    string path = $@"{this.applicationPath}\Database\Orders.xlsb";
                                                    if (!File.Exists(path))
                                                    {
                                                        return;
                                                    }

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
                                                            this.globalExportTableOptionsOpts))
                                                        {
                                                            var colFirst = 0;
                                                            var colLast = 0;

                                                            for (var colIndex = 0;
                                                                 colIndex < table.Columns.Count;
                                                                 colIndex++)
                                                            {
                                                                using (DataColumn column = table.Columns[colIndex])
                                                                {
                                                                    DateTime? dateFc = this.ulti.StringToDate(
                                                                        this.ulti.GetString(column.ColumnName));
                                                                    if (dateFc == null)
                                                                    {
                                                                        continue;
                                                                    }

                                                                    if (dateFc == dateFrom.AddDays(-maxDistance))
                                                                    {
                                                                        colFirst = colIndex;
                                                                    }

                                                                    if (dateFc != dateTo.AddDays(maxDistance))
                                                                    {
                                                                        continue;
                                                                    }

                                                                    // Once encounter dateTo, break, and record its location.
                                                                    // Further optimization, coz it IS my database.
                                                                    colLast = colIndex;
                                                                    break;
                                                                }
                                                            }

                                                            foreach (DataRow row in table.Select())
                                                            {
                                                                string productCode = this.ulti.ObjectToString(row["ProductCode"]);
                                                                string cusKeyCode = this.ulti.ObjectToString(row["CustomerKeyCode"]);

                                                                for (int colIndex = colFirst;
                                                                     colIndex <= colLast;
                                                                     colIndex++)
                                                                {
                                                                    using (DataColumn column = table.Columns[colIndex])
                                                                    {
                                                                        // First check point. Is it a valid date?
                                                                        // ReSharper disable once PossibleInvalidOperationException
                                                                        // Because I'm confident about that.
                                                                        // ... it's my fucking database.
                                                                        DateTime? datePo = this.ulti.StringToDate(this.ulti.GetString(column.ColumnName));

                                                                        ////if (datePo == null ||
                                                                        ////    datePo > dateTo.AddDays(maxDistance) ||
                                                                        ////    datePo < dateFrom.AddDays(-maxDistance))
                                                                        ////{
                                                                        ////    continue;
                                                                        ////}

                                                                        // Second check point. Is it a valid forecast value?
                                                                        double value = this.ulti.ObjectToDouble(row[colIndex]);
                                                                        if (value <= 0)
                                                                        {
                                                                            continue;
                                                                        }

                                                                        // ReSharper disable once PossibleInvalidOperationException
                                                                        if (!dicPo.TryGetValue((DateTime)datePo, out Dictionary<string, Dictionary<string, Dictionary<CustomerOrder, bool>>> poRegion))
                                                                        {
                                                                            poRegion = new Dictionary<string, Dictionary<string, Dictionary<CustomerOrder, bool>>>();
                                                                            dicPo.Add(
                                                                                (DateTime)datePo,
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

                // Gonna wait for all reading tasks to finish.
                await Task.WhenAll(readTasks).ConfigureAwait(true);
                this.WriteToRichTextBoxOutput(
                    "Đọc xong Data - Phần 2.",
                    2);

                this.TryClear();

                List<DateTime> listDatePo = dicPo.Keys.ToList();

                this.WriteToRichTextBoxOutput("Here goes nothing.");

                var coordResult = new Dictionary<string, Dictionary<(DateTime DatePo, CustomerOrder Order, Guid randomId), (DateTime DateFc, SupplierForecast Supply)>>();

                foreach (string priority in new[]
                                                {
                                                    "B2B",
                                                    "VM+ VinEco Priority",
                                                    "VM+ VinEco",
                                                    "VM Priority",
                                                    "VM+",
                                                    "VM",
                                                    string.Empty
                                                })
                {
                    Coord(priority);
                }

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
                                // Coz Repeating myself is uncool.
                                // Return the dictionary of orders for selected regions.
                                // Dictionary as collectio of choice due to performance.
                                Dictionary<CustomerOrder, bool> GetOrderDictionary(DateTime date, string region)
                                {
                                    if (dicPo.TryGetValue(date, out Dictionary<string, Dictionary<string, Dictionary<CustomerOrder, bool>>> orderRegions)
                                        && orderRegions.TryGetValue(region, out Dictionary<string, Dictionary<CustomerOrder, bool>> orderRegionProducts)
                                        && orderRegionProducts.TryGetValue(productCode, out Dictionary<CustomerOrder, bool> result))
                                    {
                                        return result.Where(
                                                po => priority == string.Empty
                                                      || localCustomers[po.Key.CustomerKeyCode]
                                                          .CustomerType
                                                      == priority)
                                            .ToDictionary(po => po.Key, po => false);
                                    }

                                    return null;
                                }

                                Dictionary<CustomerOrder, bool> orderNorth = GetOrderDictionary(datePo, "MB");

                                Dictionary<CustomerOrder, bool> orderSouth = GetOrderDictionary(
                                    datePo.AddDays(-distance[("LD", "MB")] + distance[("LD", "MN")]),
                                    "MN");

                                // Same deal. Not gonna repeat myself.
                                // Return sum of total orders.
                                double SumOrder(Dictionary<CustomerOrder, bool> source)
                                {
                                    return source?.AsParallel().Sum(po => po.Key.QuantityOrder) ?? 0;
                                }

                                double sumPoNorth = SumOrder(orderNorth);
                                double sumPoSouth = SumOrder(orderSouth);

                                // Validation. If there's no order, well, skip.
                                if (sumPoNorth + sumPoSouth <= 0d)
                                {
                                    continue;
                                }

                                // Counterpart of GetOrderDictionary.
                                Dictionary<SupplierForecast, bool> GetForecastDictionary(DateTime date, string region)
                                {
                                    if (dicFc.TryGetValue(date, out Dictionary<string, Dictionary<string, Dictionary<SupplierForecast, bool>>> forecastRegions) && forecastRegions.TryGetValue(region, out Dictionary<string, Dictionary<SupplierForecast, bool>> forecastRegionProducts) && forecastRegionProducts.TryGetValue(productCode, out Dictionary<SupplierForecast, bool> result))
                                    {
                                        return result.OrderBy(fc => fc.Key.QuantityForecast)
                                            .ToDictionary(
                                                fc => fc.Key,
                                                fc => true);
                                    }

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

                                // Not gonna repeat myself vol.3
                                // Sum of total supplies.
                                double SumForecast(Dictionary<SupplierForecast, bool> source)
                                {
                                    return source?.AsParallel().Sum(fc => fc.Key.QuantityForecast) ?? 0;
                                }

                                double sumFcNorth = SumForecast(forecastNorth);
                                double sumFcMid = SumForecast(forecastMid);
                                double sumFcSouth = SumForecast(forecastSouth);

                                // Validation - If there's no supply, also skip.
                                if (sumFcNorth + sumFcMid + sumFcSouth <= 0d)
                                {
                                    continue;
                                }

                                // Todo - Implement Rate
                                // Working on this.
                                CalculateRate();

                                void CalculateRate()
                                {
                                    try
                                    {
                                        double northMissing = sumPoNorth - sumFcNorth;
                                        double southMissing = sumPoSouth - sumFcSouth;

                                        double rateNorth = (sumFcNorth + (sumFcMid * (northMissing / (northMissing + southMissing)))) / sumPoNorth;
                                        double rateSouth = (sumFcSouth + (sumFcMid * (southMissing / (northMissing + southMissing)))) / sumPoSouth;

                                        // Todo - Implement UpperLimit for Rates.
                                        rateNorth = Math.Min(rateNorth, 1);
                                        rateSouth = Math.Min(rateSouth, 1);

                                        LetsDoThis(orderNorth, forecastNorth, sumFcNorth, rateNorth, "MB");
                                        LetsDoThis(orderSouth, forecastSouth, sumFcSouth, rateSouth, "MN");

                                        void LetsDoThis(
                                            Dictionary<CustomerOrder, bool> orders,
                                            Dictionary<SupplierForecast, bool> supplies,
                                            double sumSupplies,
                                            double rate,
                                            string region)
                                        {
                                            if (orders == null || !(sumSupplies + sumFcMid >= 0))
                                            {
                                                return;
                                            }

                                            foreach (CustomerOrder customerOrder in orders.Keys.OrderByDescending(po => po.QuantityOrder).ToList())
                                            {
                                                if (forecastNorth != null)
                                                {
                                                    PairSupplyOrder(customerOrder, ref orders, ref supplies, rate, region, region);

                                                    continue;
                                                }

                                                if (forecastMid != null)
                                                {
                                                    PairSupplyOrder(customerOrder, ref orders, ref forecastMid, rate, region, "LD");
                                                }
                                            }
                                        }

                                        void PairSupplyOrder(
                                            CustomerOrder customerOrder,
                                            ref Dictionary<CustomerOrder, bool> orders,
                                            ref Dictionary<SupplierForecast, bool> forecasts,
                                            double rate,
                                            string cusRegion,
                                            string supRegion)
                                        {
                                            // Just in case.
                                            // Validation.
                                            // Why the heck is this empty in the first place?
                                            if (rate <= 0 || forecasts == null || !forecasts.Any())
                                            {
                                                return;
                                            }

                                            try
                                            {
                                                // Warning: Black magic is happening here.
                                                // Proceed with cautions.
                                                // You have been warned.
                                                forecasts = forecasts.OrderByDescending(s => s.Key.QuantityForecast)
                                                    .ToDictionary(x => x.Key, x => x.Value);

                                                SupplierForecast supply = forecasts.First().Key;
                                                SupplierForecast supplyGiven = supply;
                                                forecasts.Remove(supply);

                                                double deliQuantity = Math.Min(
                                                    customerOrder.QuantityOrder * rate,
                                                    supply.QuantityForecast);

                                                supply.QuantityForecast -= deliQuantity;
                                                supplyGiven.QuantityForecast = deliQuantity;

                                                // ReSharper disable once SuggestVarOrType_Elsewhere
                                                if (!coordResult.TryGetValue(productCode, out var dicCoord))
                                                {
                                                    dicCoord = new Dictionary<(DateTime DatePo, CustomerOrder Order, Guid randomId), (DateTime DateFc, SupplierForecast Supply)>();
                                                    coordResult.Add(productCode, dicCoord);
                                                }

                                                // Key - Date Ordered & Customer Order.
                                                // Value - Date Processed ( substracting from the already mapped
                                                // difference in days between Supplier's Region and Customer's &
                                                // the amount of supply given to fulfill the order.
                                                dicCoord.Add(
                                                    (datePo, customerOrder, Guid.NewGuid()),
                                                    (datePo.AddDays(distance[(supRegion, cusRegion)]), supplyGiven));

                                                if (supply.QuantityForecast > 0)
                                                {
                                                    forecasts.Add(supply, false);
                                                }

                                                orders.Remove(customerOrder);
                                            }
                                            catch (Exception ex)
                                            {
                                                this.WriteToRichTextBoxOutput(ex.Message);
                                                throw;
                                            }
                                        }

                                        ////if (orderNorth != null && sumFcNorth + sumFcMid >= 0)
                                        ////{
                                        ////    foreach (CustomerOrder customerOrder in orderNorth.Keys.OrderByDescending(po => po.QuantityOrder).ToList())
                                        ////    {
                                        ////        if (forecastNorth != null)
                                        ////        {
                                        ////            PairSupplyOrder(customerOrder, orderNorth, forecastNorth, rateNorth);

                                        ////            continue;
                                        ////        }

                                        ////        if (forecastMid != null)
                                        ////        {
                                        ////            PairSupplyOrder(customerOrder, orderNorth, forecastMid, rateNorth);
                                        ////        }
                                        ////    }
                                        ////}

                                        ////if (orderSouth != null && sumFcSouth + sumFcMid >= 0)
                                        ////{
                                        ////    foreach (CustomerOrder customerOrder in orderSouth.Keys.OrderByDescending(po => po.QuantityOrder).ToList())
                                        ////    {
                                        ////        if (forecastSouth != null)
                                        ////        {
                                        ////            PairSupplyOrder(customerOrder, orderSouth, forecastSouth, rateSouth);

                                        ////            continue;
                                        ////        }

                                        ////        if (forecastMid != null)
                                        ////        {
                                        ////            PairSupplyOrder(customerOrder, orderSouth, forecastMid, rateSouth);
                                        ////        }
                                        ////    }
                                        ////}
                                    }
                                    catch (Exception ex)
                                    {
                                        this.WriteToRichTextBoxOutput(ex.Message);
                                        throw;
                                    }
                                }

                                // Todo - Select Supplier for each Order.
                            }
                        }

                        this.WriteToRichTextBoxOutput(
                            $"{(priority == string.Empty ? "Còn lại" : priority)}: {Math.Round(localWatch.Elapsed.TotalSeconds, 2).ToString(CultureInfo.InvariantCulture)}s!");

                        localWatch.Stop();
                    }
                    catch (Exception ex)
                    {
                        this.WriteToRichTextBoxOutput(ex.Message);
                        throw;
                    }
                }

                // To remove unneccessary data, only needed to calculate.
                foreach (string productCode in coordResult.Keys.ToList())
                {
                    foreach (KeyValuePair<(DateTime DatePo, CustomerOrder Order, Guid randomId), (DateTime DateFc, SupplierForecast Supply)> pair in coordResult[productCode].ToList())
                    {
                        // Because these are only here to calculate the remaining available supply.
                        if (pair.Key.DatePo < dateFrom)
                        {
                            coordResult[productCode].Remove(pair.Key);
                        }
                    }

                    // How did this even happened?
                    if (coordResult[productCode].Count == 0)
                    {
                        coordResult.Remove(productCode);
                    }
                }

                DataTable tableMastahCompact = this.ToDataTableMastahCompact(coordResult, products, customers, suppliers);

                string fileName =
                    $"Mastah Compact 100% {this.ulti.DateToString(dateFrom, "dd.MM")} - {this.ulti.DateToString(dateTo, "dd.MM")} ({this.ulti.DateToString(DateTime.Now, "yyyyMMdd HH\\hmm")}).xlsb";
                string exportPath = $@"{this.applicationPath}\Output\{fileName}";

                using (var workbook = new Workbook())
                {
                    workbook.Settings.MemorySetting = MemorySetting.MemoryPreference;

                    // Mastah
                    this.ulti.OutputExcelAspose(tableMastahCompact, workbook, true, 1);

                    workbook.Worksheets.RemoveAt("sheet1");

                    workbook.CalculateFormula();
                    workbook.Save(exportPath, SaveFormat.Xlsb);
                }

                this.ulti.DeleteEvaluationSheetInterop(exportPath);

                // The final flag.
                watch.Stop();
                this.WriteToRichTextBoxOutput(
                    $"Tổng thời gian chạy: {Math.Round(watch.Elapsed.TotalSeconds, 2).ToString(CultureInfo.InvariantCulture)}s!",
                    1);
            }
            catch (Exception ex)
            {
                this.WriteToRichTextBoxOutput(ex.Message);
                throw;
            }
            finally
            {
                this.TryClear();
            }
        }

        /// <summary>
        ///     Reading Customers from database.
        /// </summary>
        /// <returns> The <see cref="Task" />. </returns>
        private Dictionary<string, Customer> ReadCustomers()
        {
            try
            {
                var customers = new Dictionary<string, Customer>();
                if (!File.Exists($@"{this.applicationPath}\Database\Customers.xlsb"))
                {
                    return null;
                }

                using (var workbook = new Workbook(
                    $@"{this.applicationPath}\Database\Customers.xlsb",
                    new LoadOptions { MemorySetting = MemorySetting.MemoryPreference }))
                {
                    Worksheet worksheet = workbook.Worksheets[0];
                    using (DataTable table = worksheet.Cells.ExportDataTable(
                        0,
                        0,
                        worksheet.Cells.MaxDataRow + 1,
                        worksheet.Cells.MaxDataColumn + 1,
                        this.globalExportTableOptionsOpts))
                    {
                        foreach (DataRow row in table.Select())
                        {
                            customers.Add(
                                this.ulti.ObjectToString(row["Key"]),
                                new Customer
                                    {
                                        CustomerKeyCode = this.ulti.ObjectToString(row["Code"]),
                                        CustomerCode = this.ulti.ObjectToString(row["Code"]),
                                        CustomerName = this.ulti.ObjectToString(row["Name"]),
                                        CustomerBigRegion = this.ulti.ObjectToString(row["Region"]),
                                        CustomerRegion = this.ulti.ObjectToString(row["SubRegion"]),
                                        Company = this.ulti.ObjectToString(row["P&L"]),
                                        CustomerType = this.ulti.ObjectToString(row["Type"])
                                    });
                        }
                    }
                }

                return customers;
            }
            catch (Exception ex)
            {
                this.WriteToRichTextBoxOutput(ex.Message);
                throw;
            }
        }

        /// <summary>
        ///     Reading Products from database.
        /// </summary>
        /// <returns> The <see cref="Task" />. </returns>
        private Dictionary<string, Product> ReadProducts()
        {
            try
            {
                // Products
                var products = new Dictionary<string, Product>();
                if (!File.Exists($@"{this.applicationPath}\Database\Products.xlsb"))
                {
                    return null;
                }

                using (var workbook = new Workbook(
                    $@"{this.applicationPath}\Database\Products.xlsb",
                    new LoadOptions { MemorySetting = MemorySetting.MemoryPreference }))
                {
                    Worksheet worksheet = workbook.Worksheets[0];
                    using (DataTable table = worksheet.Cells.ExportDataTable(
                        0,
                        0,
                        worksheet.Cells.MaxDataRow + 1,
                        worksheet.Cells.MaxDataColumn + 1,
                        this.globalExportTableOptionsOpts))
                    {
                        foreach (DataRow row in table.Select())
                        {
                            products.Add(
                                this.ulti.ObjectToString(row["ProductCode"]),
                                new Product
                                    {
                                        ProductCode = this.ulti.ObjectToString(row["ProductCode"]),
                                        ProductName = this.ulti.ObjectToString(row["ProductName"])
                                    });
                        }
                    }
                }

                return products;
            }
            catch (Exception ex)
            {
                this.WriteToRichTextBoxOutput(ex.Message);
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
                if (!File.Exists($@"{this.applicationPath}\Database\Suppliers.xlsb"))
                {
                    return null;
                }

                using (var workbook = new Workbook(
                    $@"{this.applicationPath}\Database\Suppliers.xlsb",
                    new LoadOptions { MemorySetting = MemorySetting.MemoryPreference }))
                {
                    Worksheet worksheet = workbook.Worksheets[0];
                    using (DataTable table = worksheet.Cells.ExportDataTable(
                        0,
                        0,
                        worksheet.Cells.MaxDataRow + 1,
                        worksheet.Cells.MaxDataColumn + 1,
                        this.globalExportTableOptionsOpts))
                    {
                        foreach (DataRow row in table.Select())
                        {
                            suppliers.Add(
                                this.ulti.ObjectToString(row["SupplierCode"]),
                                new Supplier
                                    {
                                        SupplierRegion = this.ulti.ObjectToString(row["SupplierRegion"]),
                                        SupplierType = this.ulti.ObjectToString(row["SupplierType"]),
                                        SupplierCode = this.ulti.ObjectToString(row["SupplierCode"]),
                                        SupplierName = this.ulti.ObjectToString(row["SupplierName"])
                                    });
                        }
                    }
                }

                return suppliers;
            }
            catch (Exception ex)
            {
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