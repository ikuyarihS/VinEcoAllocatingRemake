using System;
using MongoDB.Bson.Serialization.Attributes;

namespace VinEcoAllocatingRemake.AllocatingInventory.Models
{
    public class ProductUnitRegion
    {
        [BsonId] public Guid _id { get; set; }

        public string Region { get; set; }
        public string OrderUnitType { get; set; }
        public double OrderUnitPer { get; set; }
        public string SaleUnitType { get; set; }
        public double SaleUnitPer { get; set; }
    }
}