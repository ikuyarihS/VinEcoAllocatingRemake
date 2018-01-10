using System;
using MongoDB.Bson.Serialization.Attributes;

namespace VinEcoAllocatingRemake.AllocatingInventory.Models
{
    public class Supplier
    {
        [BsonId] public Guid _id { get; set; }

        public Guid SupplierId { get; set; }
        public string SupplierCode { get; set; }
        public string SupplierName { get; set; }
        public string SupplierType { get; set; }
        public string SupplierRegion { get; set; }
    }
}