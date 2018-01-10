using System;
using System.Collections.Generic;
using MongoDB.Bson.Serialization.Attributes;

namespace VinEcoAllocatingRemake.AllocatingInventory.Models
{
    public class ProductForecast
    {
        [BsonId] public Guid _id { get; set; }

        public Guid ProductForecastId { get; set; }
        public Guid ProductId { get; set; }
        public List<SupplierForecast> ListSupplierForecast { get; set; }
    }
}