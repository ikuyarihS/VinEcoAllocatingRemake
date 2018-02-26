// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ToDataTable.cs" company="VinEco">
//   Shirayuki 2018.
// </copyright>
// <summary>
//   Defines the AllocatingInventory type.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics.CodeAnalysis;
using VinEcoAllocatingRemake.AllocatingInventory.Models;

namespace VinEcoAllocatingRemake.AllocatingInventory
{
    #region

    #endregion

    /// <summary>
    ///     The allocating inventory.
    /// </summary>
    // ReSharper disable once StyleCop.SA1404
    [SuppressMessage("ReSharper", "ArrangeThisQualifier")]
    public partial class AllocatingInventory
    {
        /// <summary>
        ///     Exporting CoordResult to Excel Files.
        /// </summary>
        /// <param name="coordResult"> The coord Result. </param>
        /// <param name="products"> The products. </param>
        /// <param name="customers"> The customers. </param>
        /// <param name="suppliers"> The suppliers. </param>
        /// <returns> The <see cref="DataTable" />. </returns>
        private DataTable ToDataTableMastahCompact(
            Dictionary<string, Dictionary<(DateTime DatePo, CustomerOrder Order, Guid randomId), (DateTime DateFc, SupplierForecast Supply)>> coordResult,
            IReadOnlyDictionary<string, Product>                                                                                              products,
            IReadOnlyDictionary<string, Customer>                                                                                             customers,
            IReadOnlyDictionary<string, Supplier>                                                                                             suppliers)
        {
            string YesNoFromString(bool expression)
            {
                return expression ? "Yes" : "No";
            }

            try
            {
                using (var table = new DataTable { TableName = "Mastah Compact" })
                {
                    // ReSharper disable once SuggestVarOrType_SimpleTypes
                    foreach (var details in new (string columnName, Type columnType, object columnDefaultValue)[]
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
                                                    ("Nhu cầu", typeof(double), this.ulti.DoubleToObject(0)),
                                                    ("Đáp ứng", typeof(double), this.ulti.DoubleToObject(0)),
                                                    ("Nguồn", typeof(string), null),
                                                    ("Vùng sản xuất", typeof(string), null),
                                                    ("Mã NCC", typeof(string), null),
                                                    ("Tên NCC", typeof(string), null),
                                                    ("Ngày sơ chế", typeof(DateTime), null),
                                                    ("NoSup", typeof(double), this.ulti.DoubleToObject(0)),
                                                    ("KPI", typeof(double), this.ulti.DoubleToObject(0)),
                                                    ("Label", typeof(string), null),
                                                    ("CodeSFG", typeof(string), null),
                                                    ("IsNoSup", typeof(bool), this.ulti.BoolToObject(false))
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
                        foreach (KeyValuePair<(DateTime DatePo, CustomerOrder Order, Guid randomId), (DateTime DateFc, SupplierForecast Supply)> pair in coordResult[productCode])
                        {
                            Product  product  = products[productCode];
                            Customer customer = customers[pair.Key.Order.CustomerKeyCode];
                            Supplier supplier = suppliers[pair.Value.Supply.SupplierCode];

                            // Building 'unique' rowKey to identify rows.
                            string rowKey =
                                $"{this.ulti.DateToString(pair.Key.DatePo, "yyyyMMdd")}-{customer.CustomerType}-{customer.Company}-{customer.CustomerBigRegion}-{this.ulti.DateToString(pair.Value.DateFc, "yyyyMMdd")}-{supplier.SupplierCode}";

                            // Initializing
                            DataRow dr;

                            // ... And check if row exists yet
                            if (!dicRow.TryGetValue(rowKey, out int rowIndex))
                            {
                                // If not.
                                dr = table.NewRow();

                                // Coz index is less than count by 1, have to add first.
                                dicRow.Add(rowKey, table.Rows.Count);

                                // Finally, add into DataTable.
                                table.Rows.Add(dr);

                                dr["Mã 6 ký tự"]         = productCode;
                                dr["Tên sản phẩm"]       = product.ProductName;
                                dr["Nhóm sản phẩm"]      = product.ProductClassification;
                                dr["ProductOrientation"] = product.ProductOrientation;
                                dr["ProductClimate"]     = product.ProductClimate;
                                dr["ProductionGroup"]    = product.ProductionGroup;
                                ////dr["Ghi chú"] =
                                ////    product.ProductNote.Contains(
                                ////        customer.CustomerBigRegion == "Miền Nam"
                                ////            ? "South"
                                ////            : "North")
                                ////        ? "Ok"
                                ////        : "Out of List";
                                dr["Loại cửa hàng"] = customer.CustomerType;
                                dr["P&L"]           = customer.Company;

                                // dr["Ngày tiêu thụ"] = (int)(DatePO.Date - _dateBase).TotalDays + 2;
                                dr["Ngày tiêu thụ"] = pair.Key.DatePo; // this.ulti.DateToObject(pair.Key.DatePo, "yyyyMMdd");
                                dr["Vùng tiêu thụ"] = customer.CustomerBigRegion;

                                // Todo - Add YesNoSubRegion here.
                                // dr["Tỉnh tiêu thụ"] = YesNoSubRegion ? customer.CustomerRegion : null;
                                dr["Vùng SX yêu cầu"] = pair.Key.Order.DesiredRegion;
                                dr["Nguồn yêu cầu"]   = pair.Key.Order.DesiredSource;
                            }
                            else
                            {
                                // If exists.
                                dr = table.Rows[rowIndex];
                            }

                            dr["Nhu cầu"] = this.ulti.DoubleToObject((double) dr["Nhu cầu"] + pair.Key.Order.QuantityOrder);
                            dr["Đáp ứng"] = this.ulti.DoubleToObject((double) dr["Đáp ứng"] + pair.Value.Supply.QuantityForecast);

                            if (pair.Value.Supply.QuantityForecast > 0)
                            {
                                // If there's a supplier.
                                dr["Nguồn"]         = supplier.SupplierType;
                                dr["Vùng sản xuất"] = supplier.SupplierRegion;
                                dr["Mã NCC"]        = supplier.SupplierCode;
                                dr["Tên NCC"]       = supplier.SupplierName;
                                dr["Ngày sơ chế"]   = pair.Value.DateFc; // this.ulti.DateToObject(pair.Value.DateFc, "yyyyMMdd");
                                dr["Label"]         = YesNoFromString(pair.Value.Supply.LabelVinEco);
                                dr["CodeSFG"]       =
                                    $"{productCode}1{this.ulti.IntToObject((supplier.SupplierRegion == "Lâm Đồng" ? 0 : 2) + (pair.Value.Supply.LabelVinEco ? 1 : 0))}";
                            }
                            else
                            {
                                // Otherwise.
                                dr["Nguồn"] = "Không đáp ứng";
                            }
                        }
                    }

                    foreach (DataRow dr in table.Select())
                    {
                        dr["NoSup"] = this.ulti.DoubleToObject(this.ulti.ZeroIfNegative(dr["Nhu cầu"], dr["Đáp ứng"]));
                        if ((double) dr["NoSup"] > 1)
                        {
                            dr["IsNoSup"] = this.ulti.BoolToObject(true);
                        }
                    }

                    return table;
                }
            }
            catch (Exception ex)
            {
                this.WriteToRichTextBoxOutput(ex.Message);
                throw;
            }
        }
    }
}