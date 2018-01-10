using System;
using System.Collections.Generic;
using MongoDB.Bson.Serialization.Attributes;

namespace VinEcoAllocatingRemake.AllocatingInventory.Models
{
    public class CoordResult
    {
        [BsonId] public Guid _id { get; set; }

        public Guid CoordResultId { get; set; }

        [BsonDateTimeOptions(Kind = DateTimeKind.Local)]
        public DateTime DateOrder { get; set; }

        public List<CoordResultDate> ListCoordResultDate { get; set; }
    }
}