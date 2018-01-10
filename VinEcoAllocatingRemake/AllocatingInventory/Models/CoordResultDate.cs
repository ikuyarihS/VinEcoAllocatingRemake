using System;
using System.Collections.Generic;
using MongoDB.Bson.Serialization.Attributes;

namespace VinEcoAllocatingRemake.AllocatingInventory.Models
{
    public class CoordResultDate
    {
        [BsonId] public Guid _id { get; set; }

        public Guid CoordResultDateId { get; set; }
        public Guid ProductId { get; set; }
        public List<CoordinateDate> ListCoordinateDate { get; set; }
    }
}