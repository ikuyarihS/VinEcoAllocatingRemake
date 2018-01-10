using System;
using System.Collections.Generic;
using MongoDB.Bson.Serialization.Attributes;

namespace VinEcoAllocatingRemake.AllocatingInventory.Models
{
    public class PurchaseOrderDate
    {
        [BsonId] public Guid _id { get; set; }

        public Guid PurchaseOrderDateId { get; set; }

        [BsonDateTimeOptions(Kind = DateTimeKind.Local)]
        public DateTime DateOrder { get; set; }

        public List<ProductOrder> ListProductOrder { get; set; }
    }
}