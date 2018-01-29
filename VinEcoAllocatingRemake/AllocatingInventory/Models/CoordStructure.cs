#region

using System;
using System.Collections.Generic;

#endregion

namespace VinEcoAllocatingRemake.AllocatingInventory.Models
{
    #region Declaring Model

    public class CoordStructure : IDisposable
    {
        // public Dictionary<DateTime, Dictionary<Product, Dictionary<CustomerOrder, Dictionary<SupplierForecast, DateTime>>>> dicCoord;
        public Dictionary<(DateTime DatePo, Product Product, CustomerOrder Order), (DateTime DateFc, SupplierForecast
            Supply)> dicCoord;

        public Dictionary<Guid, Customer> dicCustomer;

        // public Dictionary<DateTime, Dictionary<Product, Dictionary<SupplierForecast, double>>> dicDeli;
        public Dictionary<(DateTime DateFc, string ProductCode, SupplierForecast Supply), double> dicDeli;

        // public Dictionary<DateTime, Dictionary<Product, Dictionary<SupplierForecast, bool>>> dicFC;
        public Dictionary<(DateTime DateFc, string ProductCode, string SupplierCode), (SupplierForecast Supply, bool)>
            dicFC;

        // public Dictionary<DateTime, Dictionary<Product, Dictionary<CustomerOrder, bool>>> dicPO;
        public Dictionary<(DateTime DatePo, string ProductCode, string CustomerKeyCode), (CustomerOrder Order, bool)>
            dicPO;

        public Dictionary<Guid, Product> dicProduct;

        public Dictionary<Guid, ProductCrossRegion> dicProductCrossRegion;

        public Dictionary<string, ProductRate> dicProductRate;

        public Dictionary<Guid, Supplier> dicSupplier;

        public Dictionary<string, byte> dicTransferDays;

        public CoordStructure()
        {
            dicCoord =
                new Dictionary<(DateTime DatePo, Product Product, CustomerOrder Order), (DateTime DateFc,
                    SupplierForecast Supply)>();
            dicCustomer = new Dictionary<Guid, Customer>();
            dicDeli = new Dictionary<(DateTime DateFc, string ProductCode, SupplierForecast Supply), double>();
            dicFC =
                new Dictionary<(DateTime DateFc, string ProductCode, string SupplierCode), (SupplierForecast Supply,
                    bool)>();
            dicPO =
                new Dictionary<(DateTime DatePo, string ProductCode, string CustomerKeyCode), (CustomerOrder Order, bool
                    )>();
            dicProduct = new Dictionary<Guid, Product>();
            dicProductCrossRegion = new Dictionary<Guid, ProductCrossRegion>();
            dicProductRate = new Dictionary<string, ProductRate>();
            dicSupplier = new Dictionary<Guid, Supplier>();
            dicTransferDays = new Dictionary<string, byte>();
        }

        public void Dispose()
        {
            dicPO = null;
            dicFC = null;
            dicCoord = null;
            dicSupplier = null;
            dicCustomer = null;
            dicProductCrossRegion = null;
            dicProduct = null;
            dicDeli = null;
            dicTransferDays = null;
        }
    }

    #endregion
}