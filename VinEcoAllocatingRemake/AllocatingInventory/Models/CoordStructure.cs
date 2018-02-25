using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;

namespace VinEcoAllocatingRemake.AllocatingInventory.Models
    {
        #region

        #endregion

        #region Declaring Model

        [SuppressMessage("ReSharper", "ArrangeThisQualifier")]
        public class CoordStructure : IDisposable
            {
                // public Dictionary<DateTime, Dictionary<Product, Dictionary<CustomerOrder, Dictionary<SupplierForecast, DateTime>>>> dicCoord;
                public Dictionary<(DateTime DatePo, Product Product, CustomerOrder Order), (DateTime DateFc, SupplierForecast Supply)> dicCoord;

                public Dictionary<Guid, Customer> dicCustomer;

                // public Dictionary<DateTime, Dictionary<Product, Dictionary<SupplierForecast, double>>> dicDeli;
                public Dictionary<(DateTime DateFc, string ProductCode, SupplierForecast Supply), double> dicDeli;

                // public Dictionary<DateTime, Dictionary<Product, Dictionary<SupplierForecast, bool>>> dicFC;
                public Dictionary<(DateTime DateFc, string ProductCode, string SupplierCode), (SupplierForecast Supply, bool)> dicFC;

                // public Dictionary<DateTime, Dictionary<Product, Dictionary<CustomerOrder, bool>>> dicPO;
                public Dictionary<(DateTime DatePo, string ProductCode, string CustomerKeyCode), (CustomerOrder Order, bool)> dicPO;

                public Dictionary<Guid, Product> dicProduct;

                public Dictionary<Guid, ProductCrossRegion> dicProductCrossRegion;

                public Dictionary<string, ProductRate> dicProductRate;

                public Dictionary<Guid, Supplier> dicSupplier;

                public Dictionary<string, byte> dicTransferDays;

                public CoordStructure()
                    {
                        this.dicCoord =
                            new Dictionary<(DateTime DatePo, Product Product, CustomerOrder Order), (DateTime DateFc, SupplierForecast Supply)>();
                        this.dicCustomer = new Dictionary<Guid, Customer>();
                        this.dicDeli     = new Dictionary<(DateTime DateFc, string ProductCode, SupplierForecast Supply), double>();
                        this.dicFC       =
                            new Dictionary<(DateTime DateFc, string ProductCode, string SupplierCode), (SupplierForecast Supply, bool)>();
                        this.dicPO =
                            new Dictionary<(DateTime DatePo, string ProductCode, string CustomerKeyCode), (CustomerOrder Order, bool)>();
                        this.dicProduct            = new Dictionary<Guid, Product>();
                        this.dicProductCrossRegion = new Dictionary<Guid, ProductCrossRegion>();
                        this.dicProductRate        = new Dictionary<string, ProductRate>();
                        this.dicSupplier           = new Dictionary<Guid, Supplier>();
                        this.dicTransferDays       = new Dictionary<string, byte>();
                    }

                public void Dispose()
                    {
                        this.dicPO                 = null;
                        this.dicFC                 = null;
                        this.dicCoord              = null;
                        this.dicSupplier           = null;
                        this.dicCustomer           = null;
                        this.dicProductCrossRegion = null;
                        this.dicProduct            = null;
                        this.dicDeli               = null;
                        this.dicTransferDays       = null;
                    }
            }

        #endregion
    }