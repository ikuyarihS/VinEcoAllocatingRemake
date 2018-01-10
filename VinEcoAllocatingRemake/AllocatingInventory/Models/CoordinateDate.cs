using System;
using MongoDB.Bson.Serialization.Attributes;

namespace VinEcoAllocatingRemake.AllocatingInventory.Models
{
    public class CoordinateDate
    {
        [BsonId] public Guid _id { get; set; }

        public Guid CoordinateDateId { get; set; }
        public Guid CustomerOrderId { get; set; }
        public Guid? SupplierOrderId { get; set; }

        [BsonDateTimeOptions(Kind = DateTimeKind.Local)]
        public DateTime? DateDelier { get; set; }
    }
}