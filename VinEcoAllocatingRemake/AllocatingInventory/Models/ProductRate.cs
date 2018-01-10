using System;
using MongoDB.Bson.Serialization.Attributes;

namespace VinEcoAllocatingRemake.AllocatingInventory.Models
{
    public class ProductRate
    {
        [BsonId] public Guid _id { get; set; }

        public Guid ProductId { get; set; }
        public string ProductCode { get; set; }
        public double ToNorth { get; set; }
        public double ToSouth { get; set; }
    }
}