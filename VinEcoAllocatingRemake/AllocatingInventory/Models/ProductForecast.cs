using System;
using System.Collections.Generic;

namespace VinEcoAllocatingRemake.AllocatingInventory.Models
{
    public class ProductForecast
    {
        public Guid _id { get; set; }

        public Guid ProductForecastId { get; set; }
        public Guid ProductId { get; set; }
        public List<SupplierForecast> ListSupplierForecast { get; set; }
    }
}