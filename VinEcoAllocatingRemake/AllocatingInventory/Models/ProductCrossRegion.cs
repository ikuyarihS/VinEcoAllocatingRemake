using System;
using MongoDB.Bson.Serialization.Attributes;

namespace VinEcoAllocatingRemake.AllocatingInventory.Models
{
    public class ProductCrossRegion
    {
        public ProductCrossRegion()
        {
            ToNorth = true;
            ToSouth = true;
        }

        [BsonId] public Guid _id { get; set; }

        public Guid ProductId { get; set; }
        public bool ToNorth { get; set; }
        public bool ToSouth { get; set; }
    }
}