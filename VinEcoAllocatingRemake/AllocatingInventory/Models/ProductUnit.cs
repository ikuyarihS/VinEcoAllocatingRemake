using System;
using System.Collections.Generic;
using MongoDB.Bson.Serialization.Attributes;

namespace VinEcoAllocatingRemake.AllocatingInventory.Models
{
    public class ProductUnit
    {
        [BsonId] public Guid _id { get; set; }

        public Guid ProductId { get; set; }
        public string ProductCode { get; set; }
        public List<ProductUnitRegion> ListRegion { get; set; }
    }
}